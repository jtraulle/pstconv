package pt.cjmach.pstconv.mail;

import org.junit.jupiter.api.io.TempDir;
import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Properties;
import javax.mail.Session;
import javax.mail.internet.MimeMessage;
import javax.mail.MessagingException;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

public class MaildirFolderTest {

    @TempDir
    Path tempDir;

    @Test
    public void testGetMaildirFileNameMissingHeader() throws MessagingException {
        Session session = Session.getDefaultInstance(new Properties());
        MimeMessage msg = new MimeMessage(session);
        // X-PST-Delivery-Time header is missing
        
        assertThrows(NullPointerException.class, () -> {
            MaildirFolder.getMaildirFileName(msg, "123");
        });
    }

    @Test
    public void testGetMaildirFileNameEmptyHeader() throws MessagingException {
        Session session = Session.getDefaultInstance(new Properties());
        MimeMessage msg = new MimeMessage(session);
        msg.setHeader("X-PST-Delivery-Time", ""); // Header is present but empty
        
        // This might fail with NumberFormatException because Long.parseLong("") fails
        assertThrows(NumberFormatException.class, () -> {
            MaildirFolder.getMaildirFileName(msg, "123");
        });
    }

    @Test
    public void testAppendMessageWithMbsyncRenamedFile() throws MessagingException, IOException {
        File maildir = tempDir.resolve("maildir").toFile();
        MaildirStore store = new MaildirStore(Session.getDefaultInstance(new Properties()), maildir);
        MaildirFolder folder = (MaildirFolder) store.getDefaultFolder();
        folder.create(MaildirFolder.HOLDS_MESSAGES);

        Session session = Session.getDefaultInstance(new Properties());
        MimeMessage msg = new MimeMessage(session);
        msg.setHeader("X-PST-Delivery-Time", "1759841757000");
        msg.setHeader("X-Outlook-Descriptor-Id", "13039524");
        msg.setFlags(new javax.mail.Flags(javax.mail.Flags.Flag.SEEN), true);
        msg.setText("Original Content");

        // 1. First append
        folder.appendMessage(msg);

        File curDir = new File(maildir, "cur");
        File[] files = curDir.listFiles();
        assertEquals(1, files.length);
        String originalFileName = files[0].getName();
        assertEquals("1759841757000.13039524:2,S", originalFileName);

        // 2. Simulate mbsync renaming the file
        String mbsyncFileName = "1759841757000.13039524,U=10:2,S";
        File mbsyncFile = new File(curDir, mbsyncFileName);
        assertTrue(files[0].renameTo(mbsyncFile));

        // 3. Second append with the same message (different content to verify replacement)
        msg.setText("Updated Content");
        folder.appendMessage(msg);

        // Check that there is still only one file and it is the mbsync one
        files = curDir.listFiles();
        assertEquals(1, files.length, "There should be only one file (no duplicate)");
        assertEquals(mbsyncFileName, files[0].getName());

        // 4. Third append with flag change (should not overwrite mbsync file because flags are part of uniqueness according to getMaildirFileName)
        msg.setFlags(new javax.mail.Flags(javax.mail.Flags.Flag.ANSWERED), true);
        folder.appendMessage(msg);

        files = curDir.listFiles();
        assertEquals(2, files.length, "A flag change should create a new file because getMaildirFileName changes");
    }
}
