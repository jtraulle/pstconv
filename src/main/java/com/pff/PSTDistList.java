/**
 * PATCHED VERSION of PSTDistList from java-libpst
 *
 * This class shadows the original com.pff.PSTDistList to fix critical bugs:
 * - Bug #1: GUID copy error (copying to wrong variable)
 * - Bug #2: Operator precedence error in entryAddressType calculation
 * - Bug #3: No error handling for corrupted members
 * - Bug #4: No array bounds validation
 *
 * Original library: java-libpst 0.9.3
 * Date: 2025-02-01
 *
 * TODO: Remove this patch when java-libpst releases a fix
 */
package com.pff;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

/**
 * PST DistList for extracting Addresses from Distribution lists.
 *
 * @author Richard Johnson
 */
public class PSTDistList extends PSTMessage {

    /**
     * constructor.
     *
     * @param theFile
     *            pst file
     * @param descriptorIndexNode
     *            index of the list
     * @throws PSTException
     *             on parsing error
     * @throws IOException
     *             on data access error
     */
    PSTDistList(final PSTFile theFile, final DescriptorIndexNode descriptorIndexNode) throws PSTException, IOException {
        super(theFile, descriptorIndexNode);
    }

    /**
     * Internal constructor for performance.
     *
     * @param theFile
     *            pst file
     * @param folderIndexNode
     *            index of the list
     * @param table
     *            the PSTTableBC this object is represented by
     * @param localDescriptorItems
     *            additional external items that represent
     *            this object.
     */
    PSTDistList(final PSTFile theFile, final DescriptorIndexNode folderIndexNode, final PSTTableBC table,
                final HashMap<Integer, PSTDescriptorItem> localDescriptorItems) {
        super(theFile, folderIndexNode, table, localDescriptorItems);
    }

    /**
     * Find the next two null bytes in an array given start.
     *
     * @param data
     *            the array to search
     * @param start
     *            the starting index
     * @return position of the next null char
     */
    private int findNextNullChar(final byte[] data, int start) {
        for (; start < data.length - 1; start += 2) {
            if (data[start] == 0 && data[start + 1] == 0) {
                break;
            }
        }
        return start;
    }

    /**
     * identifier for one-off entries.
     */
    private final byte[] oneOffEntryIdUid = { (byte) 0x81, (byte) 0x2b, (byte) 0x1f, (byte) 0xa4, (byte) 0xbe,
            (byte) 0xa3, (byte) 0x10, (byte) 0x19, (byte) 0x9d, (byte) 0x6e, (byte) 0x00, (byte) 0xdd, (byte) 0x01,
            (byte) 0x0f, (byte) 0x54, (byte) 0x02 };

    /**
     * identifier for wrapped entries.
     */
    private final byte[] wrappedEntryIdUid = { (byte) 0xc0, (byte) 0x91, (byte) 0xad, (byte) 0xd3, (byte) 0x51,
            (byte) 0x9d, (byte) 0xcf, (byte) 0x11, (byte) 0xa4, (byte) 0xa9, (byte) 0x00, (byte) 0xaa, (byte) 0x00,
            (byte) 0x47, (byte) 0xfa, (byte) 0xa4 };

    /**
     * Inner class to represent distribution list one-off entries.
     */
    public class OneOffEntry {
        /** display name. */
        private String displayName = "";

        /**
         * @return display name
         */
        public String getDisplayName() {
            return this.displayName;
        }

        /** address type (smtp). */
        private String addressType = "";

        /**
         * @return address type
         */
        public String getAddressType() {
            return this.addressType;
        }

        /** email address. */
        private String emailAddress = "";

        /**
         * @return email address.
         */
        public String getEmailAddress() {
            return this.emailAddress;
        }

        /** ending position of this object in the data array. */
        private int pos = 0;

        /**
         * @return formatted record
         */
        @Override
        public String toString() {
            return String.format("Display Name: %s\n" + "Address Type: %s\n" + "Email Address: %s\n", this.displayName,
                    this.addressType, this.emailAddress);
        }
    }

    /**
     * Parse a one-off entry from this Distribution List.
     *
     * @param data
     *            the item data
     * @param pos
     *            the current position in the data.
     * @throws IOException
     *             on string reading fail
     * @return the one-off entry
     */
    private OneOffEntry parseOneOffEntry(final byte[] data, int pos) throws IOException {
        // Validation: vérifier qu'il y a assez de données
        if (pos + 4 > data.length) {
            throw new IOException("Not enough data to parse one-off entry");
        }

        final int version = (int) PSTObject.convertLittleEndianBytesToLong(data, pos, pos + 2);
        pos += 2;

        // http://msdn.microsoft.com/en-us/library/ee202811(v=exchg.80).aspx
        final int additionalFlags = (int) PSTObject.convertLittleEndianBytesToLong(data, pos, pos + 2);
        pos += 2;

        final int pad = additionalFlags & 0x8000;
        final int mae = additionalFlags & 0x0C00;
        final int format = additionalFlags & 0x1E00;
        final int m = additionalFlags & 0x0100;
        final int u = additionalFlags & 0x0080;
        final int r = additionalFlags & 0x0060;
        final int l = additionalFlags & 0x0010;
        final int pad2 = additionalFlags & 0x000F;

        int stringEnd = this.findNextNullChar(data, pos);
        if (stringEnd >= data.length) {
            throw new IOException("Invalid string termination for display name");
        }
        final byte[] displayNameBytes = new byte[stringEnd - pos];
        System.arraycopy(data, pos, displayNameBytes, 0, displayNameBytes.length);
        final String displayName = new String(displayNameBytes, "UTF-16LE");
        pos = stringEnd + 2;

        if (pos >= data.length) {
            throw new IOException("Unexpected end of data after display name");
        }

        stringEnd = this.findNextNullChar(data, pos);
        if (stringEnd >= data.length) {
            throw new IOException("Invalid string termination for address type");
        }
        final byte[] addressTypeBytes = new byte[stringEnd - pos];
        System.arraycopy(data, pos, addressTypeBytes, 0, addressTypeBytes.length);
        final String addressType = new String(addressTypeBytes, "UTF-16LE");
        pos = stringEnd + 2;

        if (pos >= data.length) {
            throw new IOException("Unexpected end of data after address type");
        }

        stringEnd = this.findNextNullChar(data, pos);
        if (stringEnd >= data.length) {
            throw new IOException("Invalid string termination for email address");
        }
        final byte[] emailAddressBytes = new byte[stringEnd - pos];
        System.arraycopy(data, pos, emailAddressBytes, 0, emailAddressBytes.length);
        final String emailAddress = new String(emailAddressBytes, "UTF-16LE");
        pos = stringEnd + 2;

        final OneOffEntry out = new OneOffEntry();
        out.displayName = displayName;
        out.addressType = addressType;
        out.emailAddress = emailAddress;
        out.pos = pos;
        return out;
    }

    /**
     * Get an array of the members in this distribution list.
     *
     * @throws PSTException
     *             on corrupted data
     * @throws IOException
     *             on bad string reading
     * @return array of entries that can either be PSTDistList.OneOffEntry
     *         or a PSTObject, generally PSTContact.
     */
    public Object[] getDistributionListMembers() throws PSTException, IOException {
        final PSTTableBCItem item = this.items.get(this.pstFile.getNameToIdMapItem(0x8055, PSTFile.PSETID_Address));

        if (item == null || item.data == null) {
            return new Object[0];
        }

        // Validation: vérifier qu'il y a assez de données pour lire le count
        if (item.data.length < 8) {
            System.err.println("Warning: Distribution list data too short, expected at least 8 bytes, got "
                    + item.data.length);
            return new Object[0];
        }

        int pos = 0;
        final int count = (int) PSTObject.convertLittleEndianBytesToLong(item.data, pos, pos + 4);

        // Validation: vérifier que le count est raisonnable
        if (count < 0 || count > 10000) {
            throw new PSTException("Invalid member count: " + count);
        }

        if (count == 0) {
            return new Object[0];
        }

        pos += 4;
        pos = (int) PSTObject.convertLittleEndianBytesToLong(item.data, pos, pos + 4);

        // Validation: vérifier que le nouveau pos est dans les limites
        if (pos < 0 || pos >= item.data.length) {
            throw new PSTException("Invalid data offset: " + pos);
        }

        // Utiliser une liste pour collecter seulement les membres valides
        final List<Object> validMembers = new ArrayList<>(count);

        for (int x = 0; x < count; x++) {
            try {
                // Validation: vérifier qu'il y a assez de données pour lire l'entrée
                if (pos + 20 > item.data.length) {
                    System.err.println("Warning: Not enough data for member " + x + " at position " + pos);
                    break;
                }

                // http://msdn.microsoft.com/en-us/library/ee218661(v=exchg.80).aspx
                // http://msdn.microsoft.com/en-us/library/ee200559(v=exchg.80).aspx
                final int flags = (int) PSTObject.convertLittleEndianBytesToLong(item.data, pos, pos + 4);
                pos += 4;

                final byte[] guid = new byte[16];
                System.arraycopy(item.data, pos, guid, 0, guid.length);
                pos += 16;

                if (Arrays.equals(guid, this.wrappedEntryIdUid)) {
                    // Validation: vérifier qu'il y a assez de données pour un wrapped entry
                    if (pos + 24 > item.data.length) {
                        System.err.println("Warning: Not enough data for wrapped entry at position " + pos);
                        break;
                    }

                    /* c3 */
                    final int entryType = item.data[pos] & 0x0F;
                    final int entryAddressType = (item.data[pos] & 0x70) >> 4;  // CORRECTION: ajout de parenthèses
                    final boolean isOneOffEntryId = (item.data[pos] & 0x80) > 0;
                    pos++;
                    final int wrappedflags = (int) PSTObject.convertLittleEndianBytesToLong(item.data, pos, pos + 4);
                    pos += 4;

                    final byte[] guid2 = new byte[16];
                    System.arraycopy(item.data, pos, guid2, 0, guid2.length);  // CORRECTION: copier dans guid2 au lieu de guid
                    pos += 16;

                    final int descriptorIndex = (int) PSTObject.convertLittleEndianBytesToLong(item.data, pos, pos + 3);
                    pos += 3;

                    final byte empty = item.data[pos];
                    pos++;

                    try {
                        final PSTObject member = PSTObject.detectAndLoadPSTObject(this.pstFile, descriptorIndex);
                        validMembers.add(member);
                    } catch (PSTException e) {
                        System.err.println("Warning: Unable to load member " + x + " with descriptor index "
                                + descriptorIndex + ": " + e.getMessage());
                        // Continue avec les autres membres au lieu de tout arrêter
                    }

                } else if (Arrays.equals(guid, this.oneOffEntryIdUid)) {
                    try {
                        final OneOffEntry entry = this.parseOneOffEntry(item.data, pos);
                        pos = entry.pos;
                        validMembers.add(entry);
                    } catch (IOException e) {
                        System.err.println("Warning: Unable to parse one-off entry for member " + x + ": "
                                + e.getMessage());
                        // Arrêter le parsing car on ne peut pas déterminer la position suivante
                        break;
                    }
                } else {
                    // GUID inconnu - logger et continuer
                    System.err.println("Warning: Unknown GUID for member " + x + ", skipping");
                    // On ne peut pas déterminer la taille de cette entrée, on arrête
                    break;
                }

            } catch (Exception e) {
                System.err.println("Error processing member " + x + ": " + e.getMessage());
                // Continue avec les autres membres si possible
            }
        }

        return validMembers.toArray(new Object[0]);
    }

    /**
     * Get an array of the members in this distribution list, returning only valid members.
     * This method never throws exceptions and returns an empty array if there's any error.
     *
     * @return array of valid entries that can either be PSTDistList.OneOffEntry
     *         or a PSTObject, generally PSTContact. Never null.
     */
    public Object[] getDistributionListMembersSafe() {
        try {
            return getDistributionListMembers();
        } catch (PSTException | IOException e) {
            System.err.println("Error reading distribution list members: " + e.getMessage());
            return new Object[0];
        }
    }
}