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

import com.pff.PSTException;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.util.Collections;
import javax.mail.Address;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Store;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;
import org.junit.jupiter.api.BeforeEach;

/**
 *
 * @author cmachado
 */
public class PstConverterTest {
    PstConverter instance;
    
    @BeforeEach
    public void setUp() {
        instance = new PstConverter();
    }
    
    @AfterEach
    public void tearDown() {
        instance = null;
    }
    
    private void testConvertInputFile(MailMessageFormat format) {
        File inputFile = new File("src/test/resources/pt/cjmach/pstconv/outlook.pst");
        File outputDirectory = new File("mailbox");
        String encoding = StandardCharsets.ISO_8859_1.name();
        int expectedTotalMessageCount = 3;
        int expectedInboxMessageCount = 2; // expected count on the inbox folder, excluding child folders.
        Store store = null;
        
        try {
            PstConvertResult result = instance.convert(inputFile, outputDirectory, format, encoding);
            assertEquals(expectedTotalMessageCount, result.getMessageCount(), "Unexpected number of converted messages.");
            
            store = instance.createStore(outputDirectory, format, encoding);
            store.connect();
            
            // Root Folder / Inbox (in portuguese)
            Folder inbox = store.getFolder("Início do ficheiro de dados do Outlook").getFolder("Caixa de Entrada");
            inbox.open(Folder.READ_ONLY);
            
            Message[] messages = inbox.getMessages();
            assertEquals(expectedInboxMessageCount, messages.length, "Unexpected number of messages in inbox.");
            
            Message messageFromAbcd = null;
            for (Message message : messages) {
                Address[] from = message.getFrom();
                assertEquals(1, from.length);
                if ("abcd@as.pt".equals(from[0].toString())) {
                    messageFromAbcd = message;
                    break;
                }
            }
            assertNotNull(messageFromAbcd, "Message from abcd@as.pt not found on inbox.");
            
            String descriptorIdHeader = ((MimeMessage) messageFromAbcd).getHeader(PstConverter.DESCRIPTOR_ID_HEADER, null);
            assertEquals("2097252", descriptorIdHeader);
            
            MimeMessage mimeMsg = (MimeMessage) messageFromAbcd;
            Object content = mimeMsg.getContent();
            MimeMultipart relatedMultiPart;
            if (content instanceof MimeMultipart) {
                MimeMultipart multiPart = (MimeMultipart) content;
                String rootContentType = multiPart.getContentType().toLowerCase();
                if (rootContentType.contains("mixed")) {
                    MimeBodyPart relatedBodyPart = (MimeBodyPart) multiPart.getBodyPart(0);
                    Object relatedContent = relatedBodyPart.getContent();
                    if (relatedContent instanceof MimeMultipart) {
                        relatedMultiPart = (MimeMultipart) relatedContent;
                    } else {
                        // In case there are no inline images, it's just the alternative body part
                        MimeMultipart alternativeMultiPart = (MimeMultipart) relatedContent;
                        try (InputStream stream = alternativeMultiPart.getBodyPart(0).getInputStream()) {
                            String bodyContent = IOUtils.toString(stream, StandardCharsets.US_ASCII.name());
                            assertEquals("Teste 23:34", bodyContent);
                        }
                        return;
                    }
                } else if (rootContentType.contains("related")) {
                    relatedMultiPart = multiPart;
                } else if (rootContentType.contains("alternative")) {
                    try (InputStream stream = multiPart.getBodyPart(0).getInputStream()) {
                        String bodyContent = IOUtils.toString(stream, StandardCharsets.US_ASCII.name());
                        assertEquals("Teste 23:34", bodyContent);
                    }
                    return;
                } else {
                    fail("Unexpected multipart type: " + multiPart.getContentType());
                    return;
                }
            } else if (content instanceof String) {
                assertEquals("Teste 23:34", content);
                return;
            } else {
                fail("Unexpected content type: " + content.getClass().getName());
                return;
            }
            MimeBodyPart alternativeBodyPart = (MimeBodyPart) relatedMultiPart.getBodyPart(0);
            MimeMultipart alternativeMultiPart = (MimeMultipart) alternativeBodyPart.getContent();
            try (InputStream stream = alternativeMultiPart.getBodyPart(0).getInputStream()) {
                String bodyContent = IOUtils.toString(stream, StandardCharsets.US_ASCII.name());
                assertEquals("Teste 23:34", bodyContent);
            }
        } catch (Exception ex) {
            fail(ex);
        } finally {
            if (store != null) {
                try {
                    store.close();
                } catch (MessagingException ignore) {}
            }
            try {
                FileUtils.deleteDirectory(outputDirectory);
            } catch (IOException ignore) { }
        }
    }
    
    @Test
    public void testConvertInputFileSuccess() {
        testConvertInputFile(MailMessageFormat.EML);
        testConvertInputFile(MailMessageFormat.MBOX);
        testConvertInputFile(MailMessageFormat.MAILDIR);
    }

    @Test
    public void testConvertSkipEmptyFolders() {
        File inputFile = new File("src/test/resources/pt/cjmach/pstconv/outlook.pst");
        File outputDirectory = new File("mailbox-skip-empty");
        MailMessageFormat format = MailMessageFormat.EML;
        String encoding = StandardCharsets.ISO_8859_1.name();
        
        try {
            // First convert without skipping empty folders to see what we have
            PstConvertResult resultNormal = instance.convert(inputFile, outputDirectory, format, encoding, false);
            
            Store storeNormal = instance.createStore(outputDirectory, format, encoding);
            storeNormal.connect();
            Folder rootNormal = storeNormal.getDefaultFolder();
            int folderCountNormal = countFolders(rootNormal);
            storeNormal.close();
            FileUtils.deleteDirectory(outputDirectory);

            // Then convert with skipping empty folders
            PstConvertResult resultSkip = instance.convert(inputFile, outputDirectory, format, encoding, true);
            
            Store storeSkip = instance.createStore(outputDirectory, format, encoding);
            storeSkip.connect();
            Folder rootSkip = storeSkip.getDefaultFolder();
            int folderCountSkip = countFolders(rootSkip);
            storeSkip.close();
            
            assertTrue(folderCountSkip < folderCountNormal, "Expected fewer folders when skipping empty ones. Normal: " + folderCountNormal + ", Skip: " + folderCountSkip);
            // In outlook.pst, there are some empty folders like "Contactos", "Calendário", etc.
            // Let's check specifically for "Caixa de Entrada" which should exist.
            Store storeVerify = instance.createStore(outputDirectory, format, encoding);
            storeVerify.connect();
            Folder rootFolder = storeVerify.getFolder("Início do ficheiro de dados do Outlook");
            Folder inbox = rootFolder.getFolder("Caixa de Entrada");
            assertTrue(inbox.exists(), "Inbox should exist as it contains messages.");
            storeVerify.close();

        } catch (Exception ex) {
            fail(ex);
        } finally {
            try {
                FileUtils.deleteDirectory(outputDirectory);
            } catch (IOException ignore) { }
        }
    }

    @Test
    public void testConvertMaildirTimestampAndModificationDate() {
        File inputFile = new File("src/test/resources/pt/cjmach/pstconv/outlook.pst");
        File outputDirectory = new File("target/test-maildir-timestamp");
        MailMessageFormat format = MailMessageFormat.MAILDIR;
        String encoding = StandardCharsets.ISO_8859_1.name();
        
        try {
            PstConvertResult result = instance.convert(inputFile, outputDirectory, format, encoding);
            assertTrue(result.getMessageCount() > 0);
            
            // Check file names and modification dates in the output directory
            // We know the structure from other tests: target/test-maildir-timestamp/Início do ficheiro de données do Outlook/Caixa de Entrada/cur/
            File inboxCurDir = new File(outputDirectory, "Início do ficheiro de données do Outlook/Caixa de Entrada/cur");
            if (!inboxCurDir.exists()) {
                 // Try another path if the above is wrong (depends on how normalizeString works)
                 inboxCurDir = new File(outputDirectory, "Início do ficheiro de dados do Outlook/Caixa de Entrada/cur");
            }
            
            assertTrue(inboxCurDir.exists(), "Inbox cur directory not found: " + inboxCurDir.getAbsolutePath());
            File[] files = inboxCurDir.listFiles();
            assertNotNull(files);
            assertTrue(files.length > 0);
            
            for (File f : files) {
                String name = f.getName();
                long timestampFromName = Long.parseLong(name.split("\\.")[0]);
                long lastModified = f.lastModified();
                
                // The timestamp in the name should be the same as the last modified date
                assertEquals(timestampFromName, lastModified, "File " + name + " has inconsistent timestamp and modification date.");
                
                // And it should not be "now" (approximately)
                long now = System.currentTimeMillis();
                assertTrue(now - lastModified > 1000, "File " + name + " seems to have current time instead of delivery time.");
            }
            
        } catch (Exception ex) {
            fail(ex);
        } finally {
            try {
                FileUtils.deleteDirectory(outputDirectory);
            } catch (IOException ignore) { }
        }
    }

    private int countFolders(Folder folder) throws MessagingException {
        int count = 1;
        for (Folder subFolder : folder.list()) {
            count += countFolders(subFolder);
        }
        return count;
    }

    /**
     * Test of convert method, of class PstConverter.
     */
    @Test
    public void testConvertInputFileNotFound() {
        String fileName = "/file/not/found.pst";
        File inputFile = new File(fileName);
        File outputDirectory = new File(".");
        MailMessageFormat format = MailMessageFormat.EML;
        String encoding = "UTF-8";
        FileNotFoundException ex = assertThrows(FileNotFoundException.class, () -> instance.convert(inputFile, outputDirectory, format, encoding));
        assertEquals(FileNotFoundException.class, ex.getClass());
    }
    
    @Test
    public void testConvertInputFileIllegal() {        
        File inputFile = new File("."); // invalid file
        File outputDirectory = new File(".");
        MailMessageFormat format = MailMessageFormat.EML;
        String encoding = "UTF-8";
        FileNotFoundException ex = assertThrows(FileNotFoundException.class, () -> instance.convert(inputFile, outputDirectory, format, encoding));
        assertEquals(FileNotFoundException.class, ex.getClass());
    }
    
    @Test
    public void testConvertOutputDirectoryIllegal() {
        File inputFile = new File("src/test/resources/pt/cjmach/pstconv/textfile.txt");
        File outputDirectory = new File("src/test/resources/pt/cjmach/pstconv/textfile.txt");
        MailMessageFormat format = MailMessageFormat.EML;
        String encoding = "UTF-8";
        assertThrows(PSTException.class, () -> instance.convert(inputFile, outputDirectory, format, encoding));
    }
    
    @Test
    public void testConvertOutputFormatNull() {
        File inputFile = new File("src/test/resources/pt/cjmach/pstconv/outlook.pst");
        File outputDirectory = new File(".");
        MailMessageFormat format = null;
        String encoding = "UTF-8";
        IllegalArgumentException iae = assertThrows(IllegalArgumentException.class, () -> instance.convert(inputFile, outputDirectory, format, encoding));
        assertEquals("format is null.", iae.getMessage());
    }
    
    @Test
    public void testConvertEncodingInvalid() {
        File inputFile = new File("src/test/resources/pt/cjmach/pstconv/outlook.pst");
        File outputDirectory = new File(".");
        MailMessageFormat format = MailMessageFormat.EML;
        String encoding = "invalid encoding";
        IllegalArgumentException iae = assertThrows(IllegalArgumentException.class, () -> instance.convert(inputFile, outputDirectory, format, encoding));
        assertEquals(encoding, iae.getMessage());
    }

    @Test
    public void testConvertRenameFolders() {
        File inputFile = new File("src/test/resources/pt/cjmach/pstconv/outlook.pst");
        File outputDirectory = new File("mailbox-rename");
        MailMessageFormat format = MailMessageFormat.EML;
        String encoding = StandardCharsets.ISO_8859_1.name();
        
        // Rename "Caixa de Entrada" to "Inbox"
        instance.setFolderNamesMap(Collections.singletonMap("Caixa de Entrada", "Inbox"));
        
        try {
            instance.convert(inputFile, outputDirectory, format, encoding, false);
            
            Store store = instance.createStore(outputDirectory, format, encoding);
            store.connect();
            Folder defaultFolder = store.getDefaultFolder();
            
            Folder rootFolder = null;
            for (Folder f : defaultFolder.list()) {
                if (f.getName().equals("Início do ficheiro de dados do Outlook")) {
                    rootFolder = f;
                    break;
                }
            }
            assertNotNull(rootFolder, "Root folder not found");
            
            Folder inboxRenamed = rootFolder.getFolder("Inbox");
            assertTrue(inboxRenamed.exists(), "Folder 'Caixa de Entrada' should have been renamed to 'Inbox'");
            
            Folder oldInbox = rootFolder.getFolder("Caixa de Entrada");
            assertFalse(oldInbox.exists(), "Folder 'Caixa de Entrada' should no longer exist under that name");
            
            store.close();
        } catch (Exception ex) {
            fail(ex);
        } finally {
            try {
                FileUtils.deleteDirectory(outputDirectory);
            } catch (IOException ignore) { }
        }
    }

    @Test
    public void testConvertFormatTHTXT() {
        File inputFile = new File("src/test/resources/pt/cjmach/pstconv/outlook.pst");
        File outputDirectory = null;
        MailMessageFormat format = MailMessageFormat.TH_TXT;
        String encoding = "UTF-8";
        
        PrintStream oldOut = System.out;
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (PrintStream newOut = new PrintStream(baos)) {
            System.setOut(newOut);
            PstConvertResult result = instance.convert(inputFile, outputDirectory, format, encoding, false);
            assertNotNull(result);
            assertEquals(0, result.getMessageCount());
            
            String output = baos.toString();
            // Verify that the output contains the bracketed item count.
            // In the test PST, "Caixa de Entrada" should have 2 messages.
            assertTrue(output.contains("[2] Caixa de Entrada") || output.contains("[2] Inbox"), "Output should contain item count in brackets. Output: " + output);
            // And it should contain some empty folders
            assertTrue(output.contains("[0] Contactos") || output.contains("[0] Contacts"), "Output should contain empty folders when skipEmptyFolders is false. Output: " + output);
        } catch (Exception ex) {
            fail(ex);
        } finally {
            System.setOut(oldOut);
        }
    }

    @Test
    public void testConvertFormatTHTXTSkipEmpty() {
        File inputFile = new File("src/test/resources/pt/cjmach/pstconv/outlook.pst");
        File outputDirectory = null;
        MailMessageFormat format = MailMessageFormat.TH_TXT;
        String encoding = "UTF-8";
        
        PrintStream oldOut = System.out;
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (PrintStream newOut = new PrintStream(baos)) {
            System.setOut(newOut);
            PstConvertResult result = instance.convert(inputFile, outputDirectory, format, encoding, true);
            assertNotNull(result);
            
            String output = baos.toString();
            // Should still contain non-empty folders
            assertTrue(output.contains("[2] Caixa de Entrada") || output.contains("[2] Inbox"), "Output should contain non-empty folders. Output: " + output);
            // Should NOT contain empty folders
            assertFalse(output.contains("[0] Contactos") || output.contains("[0] Contacts"), "Output should NOT contain empty folders when skipEmptyFolders is true. Output: " + output);
            assertFalse(output.contains("[0] Calendário") || output.contains("[0] Calendar"), "Output should NOT contain empty folders when skipEmptyFolders is true. Output: " + output);
        } catch (Exception ex) {
            fail(ex);
        } finally {
            System.setOut(oldOut);
        }
    }
}
