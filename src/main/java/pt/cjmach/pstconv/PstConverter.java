/*
 *  Copyright 2022-2025 Carlos Machado
 *
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 *  distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific language governing permissions and
 *  limitations under the License.
 */
package pt.cjmach.pstconv;

import pt.cjmach.pstconv.mail.EmlStore;
import pt.cjmach.pstconv.mail.MaildirFolder;
import pt.cjmach.pstconv.mail.MaildirStore;
import com.pff.PSTAppointment;
import com.pff.PSTAttachment;
import com.pff.PSTContact;
import com.pff.PSTDistList;
import com.pff.PSTTask;
import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;
import com.pff.PSTObject;
import com.pff.PSTRecipient;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.CoderResult;
import java.nio.charset.MalformedInputException;
import java.util.Collections;
import java.util.Date;
import java.util.Enumeration;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeSet;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.InternetHeaders;
import javax.mail.internet.MailDateFormat;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.util.ByteArrayDataSource;
import net.fortuna.mstor.model.MStorStore;
import org.apache.commons.lang3.time.StopWatch;
import org.apache.tika.mime.MimeTypeException;
import org.apache.tika.mime.MimeTypes;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Converts an Outlook OST/PST file to MBox or EML format.
 *
 * @author cmachado
 */
public class PstConverter {

    private static final Logger logger = LoggerFactory.getLogger(PstConverter.class);
    private static final MailDateFormat RFC822_DATE_FORMAT = new MailDateFormat();
    
    private Map<String, String> folderNamesMap = Collections.emptyMap();

    /**
     * Name of the custom header added to each converted message to allow to
     * easily trace back the original message from OST/PST file.
     */
    public static final String DESCRIPTOR_ID_HEADER = "X-Outlook-Descriptor-Id"; // NOI18N
    public static final String DELIVERY_TIME_HEADER = "X-PST-Delivery-Time"; // NOI18N

    /**
     * Default constructor.
     */
    public PstConverter() {
    }

    /**
     * Set the folder names map.
     * 
     * @param folderNamesMap 
     */
    public void setFolderNamesMap(Map<String, String> folderNamesMap) {
        if (folderNamesMap == null) {
            this.folderNamesMap = Collections.emptyMap();
        } else {
            this.folderNamesMap = folderNamesMap;
        }
    }

    Store createStore(File directory, MailMessageFormat format, String encoding) {
        switch (format) {
            case EML: {
                Properties sessionProps = new Properties(System.getProperties());
                Session session = Session.getDefaultInstance(sessionProps);
                return new EmlStore(session, directory);
            }

            case MAILDIR: {
                Properties sessionProps = new Properties(System.getProperties());
                Session session = Session.getDefaultInstance(sessionProps);
                return new MaildirStore(session, directory);
            }

            case MBOX: {
                // see: https://github.com/micronode/mstor#system-properties
                System.setProperty("mstor.mbox.metadataStrategy", "none"); // NOI18N
                System.setProperty("mstor.mbox.encoding", encoding); // NOI18N
                System.setProperty("mstor.mbox.bufferStrategy", "default"); // NOI18N
                System.setProperty("mstor.cache.disabled", "true"); // NOI18N

                Properties sessionProps = new Properties(System.getProperties());
                Session session = Session.getDefaultInstance(sessionProps);
                return new MStorStore(session, new URLName("mstor:" + directory)); // NOI18N
            }
            default:
                throw new IllegalArgumentException("Unsupported mail format: " + format);
        }
    }

    /**
     * Extracts the Outlook Descriptor ID Header value from each previously
     * converted message by this converter. This method can be used to test if a
     * PST file conversion was executed as expected, allowing the comparison
     * between the returned set of ids with the ones found on the original PST.
     *
     * @param directory Directory where to find the messages.
     * @param format The message format (MBOX or EML).
     * @param encoding The message encoding.
     * @return A set with all the found message ids.
     * @throws MessagingException
     */
    public Set<Long> extractDescriptorIds(File directory, MailMessageFormat format, String encoding) throws MessagingException {
        if (!directory.exists()) {
            throw new IllegalArgumentException(String.format("Inexistent directory: %s", directory.getAbsolutePath()));
        }
        if (format == null) {
            throw new IllegalArgumentException("format is null.");
        }
        Charset.forName(encoding); // throws UnsupportedCharsetException if encoding is invalid

        // see: https://docs.oracle.com/javaee/6/api/javax/mail/internet/package-summary.html#package_description
        System.setProperty("mail.mime.address.strict", "false"); // NOI18N
        Set<Long> result = new TreeSet<>();
        Store store = createStore(directory, format, encoding);
        try {
            store.connect();
            Folder mboxRootFolder = store.getDefaultFolder();
            extractDescriptorIds(mboxRootFolder, result);
        } finally {
            try {
                store.close();
            } catch (MessagingException ignore) {
            }
        }
        return result;
    }

    void extractDescriptorIds(Folder folder, Set<Long> ids) throws MessagingException {
        folder.open(Folder.READ_ONLY);
        try {
            for (Folder subFolder : folder.list()) {
                extractDescriptorIds(subFolder, ids);
            }
            if ((folder.getType() & Folder.HOLDS_MESSAGES) != 0) {
                try {
                    for (Message msg : folder.getMessages()) {
                        String[] headerValues = msg.getHeader(DESCRIPTOR_ID_HEADER);
                        if (headerValues != null && headerValues.length > 0) {
                            long id = Long.parseLong(headerValues[0]);
                            ids.add(id);
                        }
                    }
                } catch (MessagingException ex) {
                    logger.warn("Failed to get messages for folder " + folder.getFullName(), ex);
                }
            }
        } finally {
            if (folder.isOpen()) {
                try {
                    folder.close(false);
                } catch (MessagingException | NullPointerException ignore) {
                }
            }
        }
    }

    /**
     * Converts an Outlook OST/PST file to MBox or EML format.
     *
     * @param inputFile The input PST file.
     * @param outputDirectory The directory where the email messages are
     * extracted to and saved.
     * @param format The output format (MBOX or EML).
     * @param encoding The charset encoding to use for character data.
     * @param skipEmptyFolders Do not create empty folders.
     * @param includeFolder Limit processing to a specific subfolder path.
     * @param skipRootFolder Automatically skip the root folder of the PST hierarchy.
     * @return number of successfully converted messages and the duration of the
     * operation in milliseconds.
     *
     * @throws PSTException
     * @throws MessagingException
     * @throws IOException
     */
    public PstConvertResult convert(File inputFile, File outputDirectory, MailMessageFormat format, String encoding, boolean skipEmptyFolders, String includeFolder, boolean skipRootFolder) throws PSTException, MessagingException, IOException {
        PSTFile pstFile = new PSTFile(inputFile); // throws FileNotFoundException is file doesn't exist.
        return convert(pstFile, outputDirectory, format, encoding, skipEmptyFolders, includeFolder, skipRootFolder);
    }

    /**
     * Converts an Outlook OST/PST file to MBox or EML format.
     *
     * @param inputFile The input PST file.
     * @param outputDirectory The directory where the email messages are
     * extracted to and saved.
     * @param format The output format (MBOX or EML).
     * @param encoding The charset encoding to use for character data.
     * @param skipEmptyFolders Do not create empty folders.
     * @param includeFolder Limit processing to a specific subfolder path.
     * @return number of successfully converted messages and the duration of the
     * operation in milliseconds.
     *
     * @throws PSTException
     * @throws MessagingException
     * @throws IOException
     */
    public PstConvertResult convert(File inputFile, File outputDirectory, MailMessageFormat format, String encoding, boolean skipEmptyFolders, String includeFolder) throws PSTException, MessagingException, IOException {
        return convert(inputFile, outputDirectory, format, encoding, skipEmptyFolders, includeFolder, false);
    }

    /**
     * Converts an Outlook OST/PST file to MBox or EML format.
     *
     * @param inputFile The input PST file.
     * @param outputDirectory The directory where the email messages are
     * extracted to and saved.
     * @param format The output format (MBOX or EML).
     * @param encoding The charset encoding to use for character data.
     * @param skipEmptyFolders Do not create empty folders.
     * @return number of successfully converted messages and the duration of the
     * operation in milliseconds.
     *
     * @throws PSTException
     * @throws MessagingException
     * @throws IOException
     */
    public PstConvertResult convert(File inputFile, File outputDirectory, MailMessageFormat format, String encoding, boolean skipEmptyFolders) throws PSTException, MessagingException, IOException {
        return convert(inputFile, outputDirectory, format, encoding, skipEmptyFolders, null);
    }

    /**
     * Converts an Outlook OST/PST file to MBox or EML format.
     *
     * @param inputFile The input PST file.
     * @param outputDirectory The directory where the email messages are
     * extracted to and saved.
     * @param format The output format (MBOX or EML).
     * @param encoding The charset encoding to use for character data.
     * @return number of successfully converted messages and the duration of the
     * operation in milliseconds.
     *
     * @throws PSTException
     * @throws MessagingException
     * @throws IOException
     */
    public PstConvertResult convert(File inputFile, File outputDirectory, MailMessageFormat format, String encoding) throws PSTException, MessagingException, IOException {
        return convert(inputFile, outputDirectory, format, encoding, false);
    }

    /**
     * Converts an Outlook OST/PST file to MBox or EML format.
     *
     * @param pstFile The input PST file.
     * @param outputDirectory The directory where the email messages are
     * extracted to and saved.
     * @param format The output format (MBOX or EML).
     * @param encoding The charset encoding to use for character data.
     * @param skipEmptyFolders Do not create empty folders.
     * @param includeFolder Limit processing to a specific subfolder path.
     * @param skipRootFolder Automatically skip the root folder of the PST hierarchy.
     * @return number of successfully converted messages.
     *
     * @throws PSTException
     * @throws MessagingException
     * @throws IOException
     */
    public PstConvertResult convert(PSTFile pstFile, File outputDirectory, MailMessageFormat format, String encoding, boolean skipEmptyFolders, String includeFolder, boolean skipRootFolder) throws PSTException, MessagingException, IOException {
        if (format != MailMessageFormat.TH_TXT && outputDirectory.exists() && !outputDirectory.isDirectory()) {
            throw new IllegalArgumentException(String.format("Not a directory: %s.", outputDirectory.getAbsolutePath()));
        }
        if (format == null) {
            throw new IllegalArgumentException("format is null.");
        }

        Charset charset = Charset.forName(encoding); // throws UnsupportedCharsetException if encoding is invalid

        // see: https://docs.oracle.com/javaee/6/api/javax/mail/internet/package-summary.html#package_description
        System.setProperty("mail.mime.address.strict", "false"); // NOI18N
        long messageCount = 0;

        if (format == MailMessageFormat.TH_TXT) {
            StopWatch watch = StopWatch.createStarted();
            System.out.println(".");
            listFolders(pstFile.getRootFolder(), "", skipEmptyFolders);
            watch.stop();
            return new PstConvertResult(0, watch.getTime());
        }

        if (!outputDirectory.exists() && !outputDirectory.mkdirs()) {
            throw new IOException("Failed to create output directory " + outputDirectory.getAbsolutePath());
        }
        StopWatch watch = StopWatch.createStarted();
        Store store = createStore(outputDirectory, format, encoding);
        try {
            store.connect();
            Folder rootFolder = store.getDefaultFolder();
            PSTFolder pstRootFolder = pstFile.getRootFolder();
            if (skipRootFolder) {
                PSTFolder candidate = null;
                for (PSTFolder subFolder : pstRootFolder.getSubFolders()) {
                    String displayName = subFolder.getDisplayName();
                    if (displayName.isEmpty()) {
                        continue;
                    }
                    // Prioritize known "main" folder names
                    if ("Top of Outlook Data File".equals(displayName)
                            || "Top of Personal Folders".equals(displayName)
                            || "Début du fichier de données Outlook".equals(displayName)
                            || "Partie supérieure des Dossiers personnels".equals(displayName)
                            || "IPM_SUBTREE".equals(displayName)) {
                        candidate = subFolder;
                        break;
                    }
                    // Fallback to the first non-empty named folder if no priority match is found
                    if (candidate == null) {
                        candidate = subFolder;
                    }
                }
                if (candidate != null) {
                    pstRootFolder = candidate;
                }
            }
            if (includeFolder != null && !includeFolder.isEmpty()) {
                pstRootFolder = findSubFolder(pstRootFolder, includeFolder);
                if (pstRootFolder == null) {
                    throw new IllegalArgumentException("Folder not found: " + includeFolder);
                }
            }
            messageCount = convert(pstRootFolder, rootFolder, "\\", charset, skipEmptyFolders);
            watch.stop();
        } catch (PSTException | MessagingException | IOException ex) {
            logger.error("Failed to convert PSTFile object.", ex);
            throw ex;
        } finally {
            try {
                store.close();
            } catch (MessagingException ignore) {
                // ignore exception
            }
        }
        return new PstConvertResult(messageCount, watch.getTime());
    }

    /**
     * Converts an Outlook OST/PST file to MBox or EML format.
     *
     * @param pstFile The input PST file.
     * @param outputDirectory The directory where the email messages are
     * extracted to and saved.
     * @param format The output format (MBOX or EML).
     * @param encoding The charset encoding to use for character data.
     * @param skipEmptyFolders Do not create empty folders.
     * @param includeFolder Limit processing to a specific subfolder path.
     * @return number of successfully converted messages.
     *
     * @throws PSTException
     * @throws MessagingException
     * @throws IOException
     */
    public PstConvertResult convert(PSTFile pstFile, File outputDirectory, MailMessageFormat format, String encoding, boolean skipEmptyFolders, String includeFolder) throws PSTException, MessagingException, IOException {
        return convert(pstFile, outputDirectory, format, encoding, skipEmptyFolders, includeFolder, false);
    }

    /**
     * Finds a subfolder by path.
     *
     * @param root The root folder to start the search.
     * @param path The path to the subfolder, separated by '/' or '\'.
     * @return The found PSTFolder or null if not found.
     * @throws PSTException
     * @throws IOException
     */
    private PSTFolder findSubFolder(PSTFolder root, String path) throws PSTException, IOException {
        String[] parts = path.split("[\\\\/]");
        PSTFolder current = root;
        for (String part : parts) {
            if (part.isEmpty()) continue;
            boolean found = false;
            for (PSTFolder sub : current.getSubFolders()) {
                if (sub.getDisplayName().equalsIgnoreCase(part)) {
                    current = sub;
                    found = true;
                    break;
                }
            }
            if (!found) {
                return null;
            }
        }
        return current;
    }

    /**
     * Converts an Outlook OST/PST file to MBox or EML format.
     *
     * @param pstFile The input PST file.
     * @param outputDirectory The directory where the email messages are
     * extracted to and saved.
     * @param format The output format (MBOX or EML).
     * @param encoding The charset encoding to use for character data.
     * @param skipEmptyFolders Do not create empty folders.
     * @return number of successfully converted messages.
     *
     * @throws PSTException
     * @throws MessagingException
     * @throws IOException
     */
    public PstConvertResult convert(PSTFile pstFile, File outputDirectory, MailMessageFormat format, String encoding, boolean skipEmptyFolders) throws PSTException, MessagingException, IOException {
        return convert(pstFile, outputDirectory, format, encoding, skipEmptyFolders, null);
    }

    /**
     * Converts an Outlook OST/PST file to MBox or EML format.
     *
     * @param pstFile The input PST file.
     * @param outputDirectory The directory where the email messages are
     * extracted to and saved.
     * @param format The output format (MBOX or EML).
     * @param encoding The charset encoding to use for character data.
     * @return number of successfully converted messages.
     *
     * @throws PSTException
     * @throws MessagingException
     * @throws IOException
     */
    public PstConvertResult convert(PSTFile pstFile, File outputDirectory, MailMessageFormat format, String encoding) throws PSTException, MessagingException, IOException {
        return convert(pstFile, outputDirectory, format, encoding, false);
    }

    /**
     * Traverses all PSTFolders recursively, starting from the root PSTFolder,
     * and extracts all email messages to a javax.mail.Folder.
     *
     * @param pstFolder
     * @param mailFolder
     * @param path
     * @param charset
     * @param skipEmptyFolders
     * @return
     * @throws PSTException
     * @throws IOException
     * @throws MessagingException
     */
    long convert(PSTFolder pstFolder, Folder mailFolder, String path, Charset charset, boolean skipEmptyFolders) throws PSTException, IOException, MessagingException {
        long messageCount = 0;
        if (pstFolder.getContentCount() > 0) {
            PSTObject child = pstFolder.getNextChild();

            MimeMessage[] messages = new MimeMessage[1];
            while (child != null) {
                String errorMsg = "Failed to append message id {} to folder {}.";
                PSTMessage pstMessage = (PSTMessage) child;
                try {
                    if (pstMessage instanceof PSTContact && mailFolder instanceof MaildirFolder) {
                        exportContactToVCard((PSTContact) pstMessage, (MaildirFolder) mailFolder);
                        messageCount++;
                    } else if (pstMessage instanceof PSTDistList && mailFolder instanceof MaildirFolder) {
                        if (exportDistListToVCard((PSTDistList) pstMessage, (MaildirFolder) mailFolder)) {
                            messageCount++;
                        }
                    } else if (pstMessage instanceof PSTAppointment && mailFolder instanceof MaildirFolder) {
                        if (isMeetingInvitation(pstMessage)) {
                            messages[0] = convertToMimeMessage(pstMessage, charset);
                            mailFolder.appendMessages(messages);
                        } else {
                            exportAppointmentToICalendar((PSTAppointment) pstMessage, (MaildirFolder) mailFolder);
                        }
                        messageCount++;
                    } else if (pstMessage instanceof PSTTask && mailFolder instanceof MaildirFolder) {
                        exportTaskToICalendar((PSTTask) pstMessage, (MaildirFolder) mailFolder);
                        messageCount++;
                    } else {
                        messages[0] = convertToMimeMessage(pstMessage, charset);
                        mailFolder.appendMessages(messages);
                        messageCount++;
                    }
                } catch (MessagingException ex) {
                    // if the cause of the MessagingException is a MalformedInputException,
                    // then it was probably thrown due to the encoding set by the user on 
                    // the command line.
                    if (ex.getCause() instanceof MalformedInputException) {
                        MalformedInputException mie = (MalformedInputException) ex.getCause();
                        if (mie.getStackTrace().length > 0) {
                            String className = mie.getStackTrace()[0].getClassName();
                            // if the class that throwed the exception is CoderResult,
                            // then it was caused by an encoding/decoding error.
                            if (CoderResult.class.getName().equals(className)) {
                                errorMsg = String.format("Exception thrown caused by invalid encoding setting: %s. %s",
                                        charset.displayName(), errorMsg);
                            }
                        }
                    }
                    logger.error(errorMsg, child.getDescriptorNodeId(), mailFolder.getFullName(), ex);
                } catch (PSTException | IOException ex) {
                    // Handle other exceptions as well and move on to the next 
                    // PST message.
                    logger.error(errorMsg,
                            child.getDescriptorNodeId(), mailFolder.getFullName(), ex);
                }
                try {
                    child = pstFolder.getNextChild();
                } catch (IndexOutOfBoundsException ex) {
                    // This exception is thrown by java-libpst on more recent 
                    // versions (0.9.5-SNAPSHOT). It only happens when the PST 
                    // content is read from a stream.
                    logger.error("Index out of bounds when trying to get next child on folder {} ({}).", 
                            pstFolder.getDisplayName(), pstFolder.getDescriptorNodeId());
                    // Try to continue to the next PST folder.
                    break;
                }
            }
        }
        if (pstFolder.hasSubfolders()) {
            java.util.Vector<PSTFolder> subFolders;
            try {
                subFolders = pstFolder.getSubFolders();
            } catch (PSTException ex) {
                subFolders = new java.util.Vector<>();
            }
            for (PSTFolder pstSubFolder : subFolders) {
                if (skipEmptyFolders && !hasMessages(pstSubFolder)) {
                    continue;
                }
                String folderName = pstSubFolder.getDisplayName();
                if (folderNamesMap.containsKey(folderName)) {
                    folderName = folderNamesMap.get(folderName);
                }
                folderName = PstUtil.normalizeString(folderName);
                String subPath = path + "\\" + folderName;
                Folder mboxSubFolder = mailFolder.getFolder(folderName);
                if (!mboxSubFolder.exists()) {
                    if (!mboxSubFolder.create(Folder.HOLDS_FOLDERS | Folder.HOLDS_MESSAGES)) {
                        logger.warn("Failed to create mail sub folder {}.", subPath);
                        continue;
                    }
                }
                mboxSubFolder.open(Folder.READ_WRITE);
                messageCount += convert(pstSubFolder, mboxSubFolder, subPath, charset, skipEmptyFolders);
                mboxSubFolder.close(false);
            }
        }
        return messageCount;
    }

    /**
     * Recursively checks if a PST folder or its subfolders contain any messages.
     *
     * @param pstFolder The PST folder to check.
     * @return true if the folder or any of its subfolders contain messages, false otherwise.
     * @throws PSTException
     * @throws IOException
     */
    private boolean hasMessages(PSTFolder pstFolder) throws PSTException, IOException {
        if (pstFolder.getContentCount() > 0) {
            return true;
        }
        if (pstFolder.hasSubfolders()) {
            java.util.Vector<PSTFolder> subFolders;
            try {
                subFolders = pstFolder.getSubFolders();
            } catch (PSTException ex) {
                subFolders = new java.util.Vector<>();
            }
            for (PSTFolder subFolder : subFolders) {
                if (hasMessages(subFolder)) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Traverses all PSTFolders recursively, starting from the root PSTFolder,
     * and extracts all email messages to a javax.mail.Folder.
     *
     * @param pstFolder
     * @param mailFolder
     * @param path
     * @param charset
     * @return
     * @throws PSTException
     * @throws IOException
     * @throws MessagingException
     */
    long convert(PSTFolder pstFolder, Folder mailFolder, String path, Charset charset) throws PSTException, IOException, MessagingException {
        return convert(pstFolder, mailFolder, path, charset, false);
    }

    /**
     * Converts a PSTMessage to MimeMessage.
     *
     * @param message The PSTMessage object.
     * @param charset
     * @return A new MimeMessage object.
     * @throws MessagingException
     * @throws IOException
     * @throws PSTException
     * @see
     * <a href="https://www.independentsoft.de/jpst/tutorial/exporttomimemessages.html">Export
     * to MIME messages (.eml files)</a>
     */
    MimeMessage convertToMimeMessage(PSTMessage message, Charset charset) throws MessagingException, IOException, PSTException {
        MimeMessage mimeMessage = new MimeMessage((Session) null);

        convertMessageHeaders(message, mimeMessage, charset);
        // Add custom header to easily track the original message from OST/PST file.
        mimeMessage.addHeader(DESCRIPTOR_ID_HEADER, Long.toString(message.getDescriptorNodeId()));
        mimeMessage.addHeader(DELIVERY_TIME_HEADER, Long.toString(extractInternalDate(message).getTime()));
        
        // Add flags to MimeMessage
        if (message.isRead()) {
            mimeMessage.setFlag(javax.mail.Flags.Flag.SEEN, true);
        }
        if (message.hasReplied()) {
            mimeMessage.setFlag(javax.mail.Flags.Flag.ANSWERED, true);
        }
        if (message.hasForwarded()) {
            // There is no standard flag for forwarded in javax.mail.Flags.Flag
            // We can use a user flag or a custom header. 
            // Maildir uses 'P' for passed (forwarded).
            mimeMessage.setFlags(new javax.mail.Flags("Passed"), true);
        }
        if (message.isFlagged()) {
            mimeMessage.setFlag(javax.mail.Flags.Flag.FLAGGED, true);
        }

        MimeMultipart relatedMultipart = new MimeMultipart("related"); // NOI18N
        
        convertMessageBody(message, relatedMultipart);
        MimeMultipart rootMultipart = new MimeMultipart("mixed"); // NOI18N
        convertAttachments(message, rootMultipart, relatedMultipart);
        
        if (relatedMultipart.getCount() > 1) {
            MimeBodyPart relatedBodyPart = new MimeBodyPart();
            relatedBodyPart.setContent(relatedMultipart);
            
            if (rootMultipart.getCount() > 0) {
                rootMultipart.addBodyPart(relatedBodyPart, 0);
                mimeMessage.setContent(rootMultipart);
            } else {
                mimeMessage.setContent(relatedMultipart);
            }
        } else if (relatedMultipart.getCount() == 1) {
            BodyPart bodyPart = relatedMultipart.getBodyPart(0);
            if (rootMultipart.getCount() > 0) {
                rootMultipart.addBodyPart(bodyPart, 0);
                mimeMessage.setContent(rootMultipart);
            } else {
                mimeMessage.setContent(bodyPart.getContent(), bodyPart.getContentType());
            }
        } else {
            mimeMessage.setContent(rootMultipart);
        }
        return mimeMessage;
    }

    void convertMessageHeaders(PSTMessage message, MimeMessage mimeMessage, Charset charset) throws IOException, MessagingException, PSTException {
        String messageHeaders = message.getTransportMessageHeaders();
        if (messageHeaders != null && !messageHeaders.isEmpty()) {
            try (InputStream headersStream = new ByteArrayInputStream(messageHeaders.getBytes(charset))) {
                InternetHeaders headers = new InternetHeaders(headersStream);
                headers.removeHeader("Content-Type"); // NOI18N

                Enumeration<Header> allHeaders = headers.getAllHeaders();

                while (allHeaders.hasMoreElements()) {
                    Header header = allHeaders.nextElement();
                    mimeMessage.addHeader(header.getName(), header.getValue());
                }
                String dateHeader = mimeMessage.getHeader("Date", null); // NOI18N
                if (dateHeader == null || dateHeader.isEmpty()) {
                    mimeMessage.addHeader("Date", RFC822_DATE_FORMAT.format(message.getMessageDeliveryTime())); // NOI18N
                }
            }
        } else {
            mimeMessage.setSubject(message.getSubject());
            Date sentDate = message.getClientSubmitTime();
            if (sentDate == null) {
                mimeMessage.addHeader("Date", ""); // NOI18N
            } else {
                mimeMessage.setSentDate(sentDate);
            }

            InternetAddress fromMailbox = new InternetAddress();

            String senderEmailAddress = message.getSenderEmailAddress();
            fromMailbox.setAddress(senderEmailAddress);

            String senderName = message.getSenderName();
            if (senderName != null && !senderName.isEmpty()) {
                fromMailbox.setPersonal(senderName);
            } else {
                fromMailbox.setPersonal(senderEmailAddress);
            }

            mimeMessage.setFrom(fromMailbox);

            for (int i = 0; i < message.getNumberOfRecipients(); i++) {
                PSTRecipient recipient = message.getRecipient(i);
                switch (recipient.getRecipientType()) {
                    case PSTRecipient.MAPI_TO:
                        mimeMessage.setRecipient(Message.RecipientType.TO, new InternetAddress(recipient.getEmailAddress(), recipient.getDisplayName()));
                        break;
                    case PSTRecipient.MAPI_CC:
                        mimeMessage.setRecipient(Message.RecipientType.CC, new InternetAddress(recipient.getEmailAddress(), recipient.getDisplayName()));
                        break;
                    case PSTRecipient.MAPI_BCC:
                        mimeMessage.setRecipient(Message.RecipientType.BCC, new InternetAddress(recipient.getEmailAddress(), recipient.getDisplayName()));
                        break;
                    default:
                        break;
                }
            }
        }
    }

    void convertMessageBody(PSTMessage message, MimeMultipart relatedMultipart) throws IOException, MessagingException {
        String messageBody = message.getBody();
        String messageBodyHTML = message.getBodyHTML();

        if (messageBodyHTML != null && !messageBodyHTML.isEmpty()) {
            MimeMultipart alternativeMultipart = new MimeMultipart("alternative"); // NOI18N
            // Add plain text version if available
            if (messageBody != null && !messageBody.isEmpty()) {
                MimeBodyPart textBodyPart = new MimeBodyPart();
                textBodyPart.setText(messageBody);
                alternativeMultipart.addBodyPart(textBodyPart);
            }
            
            MimeBodyPart htmlBodyPart = new MimeBodyPart();
            htmlBodyPart.setDataHandler(new DataHandler(new ByteArrayDataSource(messageBodyHTML, "text/html"))); // NOI18N
            alternativeMultipart.addBodyPart(htmlBodyPart);
            
            MimeBodyPart alternativeBodyPart = new MimeBodyPart();
            alternativeBodyPart.setContent(alternativeMultipart);
            relatedMultipart.addBodyPart(alternativeBodyPart);
        } else if (messageBody != null && !messageBody.isEmpty()) {
            MimeBodyPart textBodyPart = new MimeBodyPart();
            textBodyPart.setText(messageBody);
            relatedMultipart.addBodyPart(textBodyPart);
        } else {
            MimeBodyPart textBodyPart = new MimeBodyPart();
            textBodyPart.setText("");
            textBodyPart.addHeaderLine("Content-Type: text/plain; charset=\"utf-8\""); // NOI18N
            textBodyPart.addHeaderLine("Content-Transfer-Encoding: quoted-printable"); // NOI18N
            relatedMultipart.addBodyPart(textBodyPart);
        }
    }

    void convertAttachments(PSTMessage message, MimeMultipart rootMultipart, MimeMultipart relatedMultipart) throws MessagingException, PSTException, IOException {
        if (message instanceof PSTAppointment && isMeetingInvitation(message)) {
            String ical = getAppointmentICalendar((PSTAppointment) message);
            MimeBodyPart calendarPart = new MimeBodyPart();
            calendarPart.setDataHandler(new DataHandler(new ByteArrayDataSource(ical, "text/calendar; method=REQUEST; charset=\"utf-8\"")));
            calendarPart.addHeader("Content-Transfer-Encoding", "7bit");
            rootMultipart.addBodyPart(calendarPart);
        }

        for (int i = 0; i < message.getNumberOfAttachments(); i++) {
            PSTAttachment attachment = message.getAttachment(i);

            if (attachment != null) {
                byte[] data = getAttachmentBytes(attachment);
                if (data == null) {
                    logger.warn("Failed to extract bytes of attachment {} from message {}.", 
                            attachment.getDescriptorNodeId(), message.getDescriptorNodeId());
                    // try to add the attachment, which may still be useful even without its contents.
                    data = new byte[0];
                }
                
                MimeBodyPart attachmentBodyPart = new MimeBodyPart();
                try {
                    String mimeTag = getAttachmentMimeTag(attachment);
                    DataSource source = new ByteArrayDataSource(data, mimeTag);
                    attachmentBodyPart.setDataHandler(new DataHandler(source));

                    String contentId = attachment.getContentId();
                    
                    String fileName = coalesce("attachment-" + attachment.getDescriptorNodeId(), // NOI18N
                            attachment.getLongFilename(), attachment.getDisplayName(), attachment.getFilename());
                    attachmentBodyPart.setFileName(fileName);

                    // Inline attachments should have a Content-ID and be part of the related multipart
                    if (contentId != null && !contentId.isEmpty()) {
                        // Ensure Content-ID is enclosed in angle brackets
                        if (!contentId.startsWith("<")) {
                            contentId = "<" + contentId + ">";
                        }
                        attachmentBodyPart.setContentID(contentId);
                        attachmentBodyPart.setDisposition(Part.INLINE);
                        relatedMultipart.addBodyPart(attachmentBodyPart);
                    } else {
                        attachmentBodyPart.setDisposition(Part.ATTACHMENT);
                        rootMultipart.addBodyPart(attachmentBodyPart);
                    }
                } catch (NullPointerException ex) {
                    logger.warn("Failed to convert attachment {} from message {}.", 
                            attachment.getDescriptorNodeId(), message.getDescriptorNodeId());
                }
            }
        }
    }

    private boolean isMeetingInvitation(PSTMessage message) {
        String messageClass = message.getMessageClass();
        return messageClass != null && (messageClass.equalsIgnoreCase("IPM.Schedule.Meeting.Request")
                || messageClass.equalsIgnoreCase("IPM.Schedule.Meeting.Canceled")
                || messageClass.equalsIgnoreCase("IPM.Schedule.Meeting.Resp.Pos")
                || messageClass.equalsIgnoreCase("IPM.Schedule.Meeting.Resp.Neg")
                || messageClass.equalsIgnoreCase("IPM.Schedule.Meeting.Resp.Tent"));
    }

    static boolean isMimeTypeKnown(String mime) {
        MimeTypes types = MimeTypes.getDefaultMimeTypes();
        try {
            types.forName(mime);
            return true;
        } catch (MimeTypeException ex) {
            logger.warn("Unknown mime type {}", mime);
            return false;
        }
    }

    /**
     * Extracts the content of the PSTAttachment.
     *
     * @param attachment
     * @return A byte array with the attachment content.
     * @throws PSTException If it's not possible to get the attachment input
     * stream.
     * @throws IOException If an error occurs when reading bytes from the input
     * stream.
     */
    static byte[] getAttachmentBytes(PSTAttachment attachment) throws PSTException, IOException {
        InputStream input;
        try {
            input = attachment.getFileInputStream();
        } catch (NullPointerException ex) {
            return null;
        }
        try {
            int nread;
            byte[] buffer = new byte[4096];
            try (ByteArrayOutputStream output = new ByteArrayOutputStream()) {
                while ((nread = input.read(buffer, 0, 4096)) != -1) {
                    output.write(buffer, 0, nread);
                }
                output.flush();
                byte[] result = output.toByteArray();
                return result;
            }
        } finally {
            try {
                input.close();
            } catch (IOException ignore) { }
        }
    }

    static String getAttachmentMimeTag(PSTAttachment attachment) {
        String mimeTag = null;
        try {
            mimeTag = attachment.getMimeTag();
        } catch (NullPointerException ignore) { }
        // mimeTag should contain a valid mime type, but sometimes it doesn't.
        // To prevent throwing exceptions when the MimeMessage is validated, the
        // mimeTag value is first checked with isMimeTypeKnown(). If it's not 
        // known, the mime type is set to 'application/octet-stream.
        if (mimeTag == null || mimeTag.isEmpty()) {
            return "application/octet-stream";
        }
        if (isMimeTypeKnown(mimeTag)) {
            return mimeTag;
        }
        return "application/octet-stream";
    }

    static String coalesce(String defaultValue, String... args) {
        for (String arg : args) {
            if (arg != null && !arg.isEmpty()) {
                return arg;
            }
        }
        return defaultValue;
    }

    /**
     * Extracts the most appropriate date from a PSTMessage to serve
     * as INTERNALDATE when migrating to Maildir.
     * Priority order:
     * 1. MessageDeliveryTime  (Exchange reception date)
     * 2. ClientSubmitTime     (send date, useful for "Sent" items)
     * 3. CreationTime         (creation date in the PST)
     * 4. Current date         (last resort)
     */
    public static Date extractInternalDate(PSTMessage message) {

        Date date;

        // 1st choice: reception date on the server
        date = message.getMessageDeliveryTime();
        if (isValidDate(date)) {
            return date;
        }

        // 2nd choice: send date by the sender
        date = message.getClientSubmitTime();
        if (isValidDate(date)) {
            return date;
        }

        // 3rd choice: message creation date in the PST
        date = message.getCreationTime();
        if (isValidDate(date)) {
            return date;
        }

        // Last resort
        return new Date();
    }

    /**
     * Checks that a date is neither null nor epoch (Date(0))
     * which is returned by getDateItem() when the data is empty.
     */
    private static boolean isValidDate(Date date) {
        return date != null && date.getTime() != 0;
    }

    private void exportContactToVCard(PSTContact contact, MaildirFolder maildirFolder) throws MessagingException {
        String descriptorIndex = Long.toString(contact.getDescriptorNodeId());
        String uid = "PST-VC-" + descriptorIndex;
        String fileName = uid + ".vcf";
        File vcfFile = new File(maildirFolder.getDirectory(), fileName);

        StringBuilder vcard = new StringBuilder();
        vcard.append("BEGIN:VCARD\r\n");
        vcard.append("VERSION:3.0\r\n");
        vcard.append("UID:").append(uid).append("\r\n");

        // Basic fields
        String fullName = contact.getDisplayName();
        if (fullName != null && !fullName.isEmpty()) {
            vcard.append("FN:").append(fullName).append("\r\n");
        }
        String firstName = contact.getGivenName();
        String lastName = contact.getSurname();
        if ((firstName != null && !firstName.isEmpty()) || (lastName != null && !lastName.isEmpty())) {
            vcard.append("N:").append(coalesce("", lastName)).append(";").append(coalesce("", firstName)).append(";;;\r\n");
        }
        String nickname = contact.getNickname();
        if (nickname != null && !nickname.isEmpty()) {
            vcard.append("NICKNAME:").append(nickname).append("\r\n");
        }
        String title = contact.getTitle();
        if (title != null && !title.isEmpty()) {
            vcard.append("TITLE:").append(title).append("\r\n");
        }
        String companyName = contact.getCompanyName();
        if (companyName != null && !companyName.isEmpty()) {
            vcard.append("ORG:").append(companyName).append("\r\n");
        }
        String email = contact.getEmail1EmailAddress();
        if (email != null && !email.isEmpty()) {
            vcard.append("EMAIL:").append(email).append("\r\n");
        }

        // Phone numbers
        String workPhone = contact.getBusinessTelephoneNumber();
        if (workPhone != null && !workPhone.isEmpty()) {
            vcard.append("TEL;TYPE=work:").append(workPhone).append("\r\n");
        }
        String homePhone = contact.getHomeTelephoneNumber();
        if (homePhone != null && !homePhone.isEmpty()) {
            vcard.append("TEL;TYPE=home:").append(homePhone).append("\r\n");
        }
        String mobilePhone = contact.getMobileTelephoneNumber();
        if (mobilePhone != null && !mobilePhone.isEmpty()) {
            vcard.append("TEL;TYPE=cell:").append(mobilePhone).append("\r\n");
        }
        String faxPhone = contact.getBusinessFaxNumber();
        if (faxPhone != null && !faxPhone.isEmpty()) {
            vcard.append("TEL;TYPE=fax:").append(faxPhone).append("\r\n");
        }

        // Addresses
        String workStreet = contact.getWorkAddressStreet();
        String workCity = contact.getWorkAddressCity();
        String workState = contact.getWorkAddressState();
        String workPostalCode = contact.getWorkAddressPostalCode();
        String workCountry = contact.getWorkAddressCountry();
        String workPoBox = contact.getWorkAddressPostOfficeBox();

        if (coalesce("", workStreet, workCity, workState, workPostalCode, workCountry, workPoBox).length() > 0) {
            vcard.append("ADR;TYPE=work:").append(coalesce("", workPoBox)).append(";")
                    .append("") // Extended address (appartement/bureau 2nd line)
                    .append(";").append(coalesce("", workStreet).replace("\n", " ").replace("\r", ""))
                    .append(";").append(coalesce("", workCity))
                    .append(";").append(coalesce("", workState))
                    .append(";").append(coalesce("", workPostalCode))
                    .append(";").append(coalesce("", workCountry))
                    .append("\r\n");
        }

        String homeStreet = contact.getHomeAddressStreet();
        String homeCity = contact.getHomeAddressCity();
        String homeState = contact.getHomeAddressStateOrProvince();
        String homePostalCode = contact.getHomeAddressPostalCode();
        String homeCountry = contact.getHomeAddressCountry();
        String homePoBox = contact.getHomeAddressPostOfficeBox();

        if (coalesce("", homeStreet, homeCity, homeState, homePostalCode, homeCountry, homePoBox).length() > 0) {
            vcard.append("ADR;TYPE=home:").append(coalesce("", homePoBox)).append(";")
                    .append("")
                    .append(";").append(coalesce("", homeStreet).replace("\n", " ").replace("\r", ""))
                    .append(";").append(coalesce("", homeCity))
                    .append(";").append(coalesce("", homeState))
                    .append(";").append(coalesce("", homePostalCode))
                    .append(";").append(coalesce("", homeCountry))
                    .append("\r\n");
        }

        vcard.append("END:VCARD\r\n");

        try (java.io.FileOutputStream fos = new java.io.FileOutputStream(vcfFile)) {
            fos.write(vcard.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8));
        } catch (IOException ex) {
            throw new MessagingException("Failed to write VCARD file: " + vcfFile.getAbsolutePath(), ex);
        }
    }

    private boolean exportDistListToVCard(PSTDistList distList, MaildirFolder maildirFolder) throws MessagingException {
        Object[] members = distList.getDistributionListMembersSafe();
        if (members == null || members.length == 0) {
            return false;
        }

        String descriptorIndex = Long.toString(distList.getDescriptorNodeId());
        String uid = "PST-VL-" + descriptorIndex;
        String fileName = uid + ".vcf";
        File vcfFile = new File(maildirFolder.getDirectory(), fileName);

        StringBuilder vlist = new StringBuilder();
        vlist.append("BEGIN:VLIST\r\n");
        vlist.append("UID:").append(uid).append(".vcf").append("\r\n");
        vlist.append("VERSION:1.0\r\n");

        for (int i = 0; i < members.length; i++) {
            try {
                Object member = members[i];
                if (member instanceof PSTContact contact) {
                    vlist.append("CARD;")
                            .append(escapeICalendar(contact.getEmail1EmailAddress()))
                            .append(";FN=")
                            .append(escapeICalendar(contact.getDisplayName()))
                            .append(":")
                            .append("PST-VC-")
                            .append(contact.getDescriptorNodeId())
                            .append(".vcf")
                            .append("\r\n");
                } else if (member instanceof PSTDistList.OneOffEntry entry) {
                    vlist.append("MEMBER;FN=")
                            .append(escapeICalendar(entry.getDisplayName()))
                            .append(":")
                            .append(escapeICalendar(entry.getEmailAddress()))
                            .append("\r\n");
                }
            } catch (Exception e) {
                logger.warn("Error processing member {} of distribution list {}", i, uid, e);
            }
        }

        String groupName = distList.getDisplayName();
        if (groupName == null || groupName.isEmpty()) {
            groupName = distList.getSubject();
        }
        if (groupName != null && !groupName.isEmpty()) {
            vlist.append("FN:").append(groupName).append("\r\n");
        }
        vlist.append("END:VLIST\r\n");

        try (java.io.FileOutputStream fos = new java.io.FileOutputStream(vcfFile)) {
            fos.write(vlist.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8));
        } catch (IOException ex) {
            throw new MessagingException("Failed to write VCARD (VLIST) file: " + vcfFile.getAbsolutePath(), ex);
        }
        return true;
    }

    private String getAppointmentICalendar(PSTAppointment appointment) {
        String descriptorIndex = Long.toString(appointment.getDescriptorNodeId());
        String uid = "PST-VE-" + descriptorIndex;

        StringBuilder ical = new StringBuilder();
        ical.append("BEGIN:VCALENDAR\r\n");
        ical.append("VERSION:2.0\r\n");
        ical.append("PRODID:-//pstconv//NONSGML v1.0//EN\r\n");
        ical.append("BEGIN:VEVENT\r\n");
        ical.append("UID:").append(uid).append("\r\n");

        java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyyMMdd'T'HHmmss'Z'");
        sdf.setTimeZone(java.util.TimeZone.getTimeZone("UTC"));
        String now = sdf.format(new Date());
        ical.append("DTSTAMP:").append(now).append("\r\n");

        Date start = appointment.getStartTime();
        if (appointment.isRecurring() && appointment.getRecurrenceBase() != null) {
            start = appointment.getRecurrenceBase();
        }
        if (start != null) {
            ical.append("DTSTART:").append(sdf.format(start)).append("\r\n");
        }
        Date end = appointment.getEndTime();
        if (end != null) {
            ical.append("DTEND:").append(sdf.format(end)).append("\r\n");
        }

        String summary = appointment.getSubject();
        if (summary != null && !summary.isEmpty()) {
            ical.append("SUMMARY:").append(escapeICalendar(summary)).append("\r\n");
        }

        String location = appointment.getLocation();
        if (location != null && !location.isEmpty()) {
            ical.append("LOCATION:").append(escapeICalendar(location)).append("\r\n");
        }

        String description = appointment.getBody();
        if (description != null && !description.isEmpty()) {
            ical.append("DESCRIPTION:").append(escapeICalendar(description)).append("\r\n");
        }

        int sensitivity = appointment.getSensitivity();
        String classification;
        switch (sensitivity) {
            case 1:
            case 2:
                classification = "PRIVATE";
                break;
            case 3:
                classification = "CONFIDENTIAL";
                break;
            case 0:
            default:
                classification = "PUBLIC";
                break;
        }
        ical.append("CLASS:").append(classification).append("\r\n");

        try {
            if (appointment.isRecurring()) {
                byte[] structure = appointment.getRecurrenceStructure();
                if (structure != null && structure.length > 0) {
                    String rrule = parseRecurrenceStructure(structure);
                    if (rrule != null) {
                        ical.append("RRULE:").append(rrule).append("\r\n");
                    }
                }
            }
        } catch (Exception ex) {
            logger.warn("Failed to get recurrence structure for appointment {}", uid, ex);
        }

        ical.append("END:VEVENT\r\n");
        ical.append("END:VCALENDAR\r\n");
        return ical.toString();
    }

    private void exportAppointmentToICalendar(PSTAppointment appointment, MaildirFolder maildirFolder) throws MessagingException {
        String descriptorIndex = Long.toString(appointment.getDescriptorNodeId());
        String uid = "PST-VE-" + descriptorIndex;
        String fileName = uid + ".ics";
        File icsFile = new File(maildirFolder.getDirectory(), fileName);

        String ical = getAppointmentICalendar(appointment);

        try (java.io.FileOutputStream fos = new java.io.FileOutputStream(icsFile)) {
            fos.write(ical.getBytes(java.nio.charset.StandardCharsets.UTF_8));
        } catch (IOException ex) {
            throw new MessagingException("Failed to write iCalendar file: " + icsFile.getAbsolutePath(), ex);
        }
    }

    private void exportTaskToICalendar(PSTTask task, MaildirFolder maildirFolder) throws MessagingException {
        String descriptorIndex = Long.toString(task.getDescriptorNodeId());
        String uid = "PST-VT-" + descriptorIndex;
        String fileName = uid + ".ics";
        File icsFile = new File(maildirFolder.getDirectory(), fileName);

        StringBuilder ical = new StringBuilder();
        ical.append("BEGIN:VCALENDAR\r\n");
        ical.append("VERSION:2.0\r\n");
        ical.append("PRODID:-//pstconv//NONSGML v1.0//EN\r\n");
        ical.append("BEGIN:VTODO\r\n");
        ical.append("UID:").append(uid).append("\r\n");

        java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyyMMdd'T'HHmmss'Z'");
        sdf.setTimeZone(java.util.TimeZone.getTimeZone("UTC"));
        String now = sdf.format(new Date());
        ical.append("DTSTAMP:").append(now).append("\r\n");

        String summary = task.getSubject();
        if (summary != null && !summary.isEmpty()) {
            ical.append("SUMMARY:").append(escapeICalendar(summary)).append("\r\n");
        }

        String description = task.getBody();
        if (description != null && !description.isEmpty()) {
            ical.append("DESCRIPTION:").append(escapeICalendar(description)).append("\r\n");
        }

        Date start = task.getTaskStartDate();
        if (start != null) {
            ical.append("DTSTART:").append(sdf.format(start)).append("\r\n");
        }

        Date due = task.getTaskDueDate();
        if (due != null) {
            ical.append("DUE:").append(sdf.format(due)).append("\r\n");
        }

        Date completedDate = task.getTaskDateCompleted();
        if (completedDate != null) {
            ical.append("COMPLETED:").append(sdf.format(completedDate)).append("\r\n");
        }

        int priority = task.getImportance();
        // Outlook importance: 0=Low, 1=Normal, 2=High
        // RFC 5545 PRIORITY: 1-5 (High), 5 (Normal), 6-9 (Low), 0 (Undefined)
        if (priority == 2) {
            ical.append("PRIORITY:1\r\n");
        } else if (priority == 0) {
            ical.append("PRIORITY:9\r\n");
        } else {
            ical.append("PRIORITY:5\r\n");
        }

        int status = task.getTaskStatus();
        // Outlook task status: 0=Not Started, 1=In Progress, 2=Complete, 3=Waiting, 4=Deferred
        // RFC 5545 STATUS: NEEDS-ACTION, IN-PROCESS, COMPLETED, CANCELLED
        switch (status) {
            case 0:
                ical.append("STATUS:NEEDS-ACTION\r\n");
                break;
            case 1:
            case 3:
            case 4:
                ical.append("STATUS:IN-PROCESS\r\n");
                break;
            case 2:
                ical.append("STATUS:COMPLETED\r\n");
                break;
        }

        double percentComplete = task.getPercentComplete();
        ical.append("PERCENT-COMPLETE:").append((int) (percentComplete * 100)).append("\r\n");

        ical.append("END:VTODO\r\n");
        ical.append("END:VCALENDAR\r\n");

        try (java.io.FileOutputStream fos = new java.io.FileOutputStream(icsFile)) {
            fos.write(ical.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8));
        } catch (IOException ex) {
            throw new MessagingException("Failed to write iCalendar file: " + icsFile.getAbsolutePath(), ex);
        }
    }

    private void listFolders(PSTFolder pstFolder, String prefix, boolean skipEmptyFolders) throws PSTException, IOException {
        java.util.Vector<PSTFolder> subFolders;
        try {
            subFolders = pstFolder.getSubFolders();
        } catch (PSTException ex) {
            // Some special folders like Search Folders might fail to list children.
            return;
        }

        if (skipEmptyFolders) {
            java.util.Vector<PSTFolder> filtered = new java.util.Vector<>();
            for (PSTFolder subFolder : subFolders) {
                if (hasMessages(subFolder)) {
                    filtered.add(subFolder);
                }
            }
            subFolders = filtered;
        }

        int count = subFolders.size();
        for (int i = 0; i < count; i++) {
            PSTFolder subFolder = subFolders.get(i);
            boolean isLast = (i == count - 1);
            String folderName = subFolder.getDisplayName();
            int itemCount = subFolder.getContentCount();
            System.out.print(prefix);
            if (isLast) {
                System.out.println("└── [" + itemCount + "] " + folderName);
                listFolders(subFolder, prefix + "    ", skipEmptyFolders);
            } else {
                System.out.println("├── [" + itemCount + "] " + folderName);
                listFolders(subFolder, prefix + "│   ", skipEmptyFolders);
            }
        }
    }

    private String parseRecurrenceStructure(byte[] data) {
        if (data == null || data.length < 22) {
            return null;
        }

        // MS-OXOCAL 2.2.1.44 PidLidRecurrencePattern
        // ReaderVersion: data[0-1]
        // WriterVersion: data[2-3]
        // RecurFrequency: data[4-5]
        int recurFrequency = (data[4] & 0xFF) | ((data[5] & 0xFF) << 8);
        // PatternType: data[6-7]
        int patternType = (data[6] & 0xFF) | ((data[7] & 0xFF) << 8);
        // CalendarType: data[8-9]
        // FirstDateTime: data[10-13]
        // Period: data[14-17]
        long period = (data[14] & 0xFFL) | ((data[15] & 0xFFL) << 8) | ((data[16] & 0xFFL) << 16) | ((data[17] & 0xFFL) << 24);
        // SlidingFlag: data[18-21]

        StringBuilder rrule = new StringBuilder();
        switch (recurFrequency) {
            case 0x200A: // Daily
                rrule.append("FREQ=DAILY");
                break;
            case 0x200B: // Weekly
                rrule.append("FREQ=WEEKLY");
                // MS-OXOCAL 2.2.1.44.1.2 Weekly Recurrence Pattern
                if (data.length >= 26) {
                    int dayMask = (data[22] & 0xFF) | ((data[23] & 0xFF) << 8) | ((data[24] & 0xFF) << 16) | ((data[25] & 0xFF) << 24);
                    String byDay = convertDayMaskToByDay(dayMask);
                    if (!byDay.isEmpty()) {
                        rrule.append(";BYDAY=").append(byDay);
                    }
                }
                break;
            case 0x200C: // Monthly
                rrule.append("FREQ=MONTHLY");
                if (patternType == 2 || patternType == 3) { // MonthPattern or MonthEndPattern
                    if (data.length >= 26) {
                        int dayOfMonth = (data[22] & 0xFF) | ((data[23] & 0xFF) << 8) | ((data[24] & 0xFF) << 16) | ((data[25] & 0xFF) << 24);
                        rrule.append(";BYMONTHDAY=").append(dayOfMonth);
                    }
                } else if (patternType == 4) { // MonthNthPattern
                    if (data.length >= 30) {
                        int dayMask = (data[22] & 0xFF) | ((data[23] & 0xFF) << 8) | ((data[24] & 0xFF) << 16) | ((data[25] & 0xFF) << 24);
                        int nth = (data[26] & 0xFF) | ((data[27] & 0xFF) << 8) | ((data[28] & 0xFF) << 16) | ((data[29] & 0xFF) << 24);
                        String byDay = convertDayMaskToByDay(dayMask);
                        if (!byDay.isEmpty()) {
                            // nth: 1 to 4, or 5 for last
                            rrule.append(";BYDAY=").append(nth == 5 ? "-1" : nth).append(byDay);
                        }
                    }
                }
                break;
            case 0x200D: // Yearly
                rrule.append("FREQ=YEARLY");
                break;
            default:
                return null;
        }

        if (period > 1) {
            rrule.append(";INTERVAL=").append(period);
        }

        // End type is further in the structure, but it depends on the pattern type which affects the offset.
        // For simplicity and safety, we focus on the frequency, interval and day masking which are most critical.
        
        return rrule.toString();
    }

    private String convertDayMaskToByDay(int dayMask) {
        // MS-OXOCAL 2.2.1.44.1.2 DayMask
        // bit 0: Sunday, bit 1: Monday, ..., bit 6: Saturday
        StringBuilder sb = new StringBuilder();
        if ((dayMask & 0x01) != 0) sb.append("SU,");
        if ((dayMask & 0x02) != 0) sb.append("MO,");
        if ((dayMask & 0x04) != 0) sb.append("TU,");
        if ((dayMask & 0x08) != 0) sb.append("WE,");
        if ((dayMask & 0x10) != 0) sb.append("TH,");
        if ((dayMask & 0x20) != 0) sb.append("FR,");
        if ((dayMask & 0x40) != 0) sb.append("SA,");
        
        if (sb.length() > 0) {
            sb.setLength(sb.length() - 1); // remove last comma
        }
        return sb.toString();
    }

    private String escapeICalendar(String text) {
        return text.replace("\\", "\\\\")
                .replace(";", "\\;")
                .replace(",", "\\,")
                .replace("\n", "\\n")
                .replace("\r", "");
    }
}
