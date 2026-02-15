/*
 *  Copyright 2026 Jean Traullé
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
package pt.cjmach.pstconv.mail;

import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.mail.Flags;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Store;
import javax.mail.internet.MimeMessage;

public class MaildirFolder extends LocalFolder {

    private static final FileFilter MAILDIR_FILE_FILTER = (File pathname) -> {
        return pathname.isFile();
    };

    private final File curDir;
    private final File tmpDir;
    private final File newDir;

    public MaildirFolder(Store store, File directory) {
        super(store, directory, MAILDIR_FILE_FILTER);
        this.curDir = new File(directory, "cur");
        this.tmpDir = new File(directory, "tmp");
        this.newDir = new File(directory, "new");
    }

    @Override
    public boolean create(int type) throws MessagingException {
        boolean created = super.create(type);
        if (created && (type & HOLDS_MESSAGES) != 0) {
            curDir.mkdirs();
            tmpDir.mkdirs();
            newDir.mkdirs();
        }
        return created;
    }

    @Override
    public void appendMessage(Message msg) throws MessagingException {
        String id = getDescriptorNodeId(msg);
        String fileName = getMaildirFileName(msg, id);
        
        // Check for existing file with same base name (ignoring mbsync UID)
        File curFile = findExistingFile(fileName);
        if (curFile == null) {
            curFile = new File(curDir, fileName);
        }
        
        // Use tmp/ for initial write, then move to cur/
        File tmpFile = new File(tmpDir, fileName);
        
        try (FileOutputStream outputStream = new FileOutputStream(tmpFile)) {
            msg.writeTo(outputStream);
        } catch (IOException ex) {
            throw new MessagingException("Failed to write to tmp file", ex);
        }
        
        if (!tmpFile.renameTo(curFile)) {
             // If renameTo fails, it might be because the destination exists.
             // Although on many systems renameTo overwrites.
             // Let's try to delete it first if it exists and rename failed.
             if (curFile.exists() && !curFile.delete()) {
                 throw new MessagingException("Failed to delete existing message file: " + curFile.getName());
             }
             if (!tmpFile.renameTo(curFile)) {
                 throw new MessagingException("Failed to move message from tmp/ to cur/");
             }
        }

        // Set the file modification time to the message delivery time, if available.
        String[] deliveryTimeHeader = msg.getHeader("X-PST-Delivery-Time");
        long timestamp = Long.parseLong(deliveryTimeHeader[0]);
        curFile.setLastModified(timestamp);
    }

    private File findExistingFile(String fileName) {
        // Expected format: <timestamp>.<id>:2,<flags>
        // mbsync format: <timestamp>.<id>,U=<uid>:2,<flags>
        int colonIndex = fileName.lastIndexOf(":2,");
        if (colonIndex == -1) {
            return null;
        }
        String base = fileName.substring(0, colonIndex);
        String flags = fileName.substring(colonIndex);
        
        File[] files = curDir.listFiles();
        if (files != null) {
            for (File file : files) {
                String existingName = file.getName();
                if (existingName.endsWith(flags)) {
                    int existingColonIndex = existingName.lastIndexOf(":2,");
                    if (existingColonIndex != -1) {
                        String existingBase = existingName.substring(0, existingColonIndex);
                        // Check if existingBase starts with base followed by "," or is equal to base
                        if (existingBase.equals(base) || existingBase.startsWith(base + ",")) {
                            return file;
                        }
                    }
                }
            }
        }
        return null;
    }

    @Override
    protected LocalFolder createInstance(Store store, File directory, FileFilter fileFilter) {
        return new MaildirFolder(store, directory);
    }

    @Override
    public Message getMessage(int msgnum) throws MessagingException {
        File msgFile = getFileEntries()[msgnum - 1];
        try (FileInputStream msgFileStream = new FileInputStream(msgFile)) {
            MimeMessage msg = new MimeMessage(null, msgFileStream);
            return msg;
        } catch (IOException ex) {
            throw new MessagingException("Failed to get message", ex);
        }
    }
    
    @Override
    public File[] getFileEntries() {
        File[] curFiles = curDir.listFiles(MAILDIR_FILE_FILTER);
        return curFiles != null ? curFiles : new File[0];
    }

    @Override
    public int getMessageCount() throws MessagingException {
        return getFileEntries().length;
    }

    public File getDirectory() {
        return directory;
    }

    static String getMaildirFileName(Message msg, String descriptorIndex) throws MessagingException {
        String[] deliveryTimeHeader = msg.getHeader("X-PST-Delivery-Time");
        long timestamp = Long.parseLong(deliveryTimeHeader[0]);
        
        StringBuilder builder = new StringBuilder();
        builder.append(timestamp).append(".");
        builder.append(descriptorIndex);
        
        // Maildir info/flags
        builder.append(":2,");
        
        javax.mail.Flags flags = msg.getFlags();
        if (flags.contains(Flags.Flag.FLAGGED)) {
            builder.append("F");
        }
        if (flags.contains("Passed")) {
            builder.append("P");
        }
        if (flags.contains(javax.mail.Flags.Flag.ANSWERED)) {
            builder.append("R");
        }
        if (flags.contains(javax.mail.Flags.Flag.SEEN)) {
            builder.append("S");
        }
        
        return builder.toString();
    }
}
