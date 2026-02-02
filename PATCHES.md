# Applied Patches

## PSTDistList.java

**Original version:** java-libpst 0.9.3  
**Modified file:** com/pff/PSTDistList.java  
**Reason:** Critical bug fixes in getDistributionListMembers()

### Fixed Bugs:
1. Copy into wrong GUID variable (guid instead of guid2)
2. Incorrect operator precedence for entryAddressType
3. Missing error handling for corrupted members
4. No array bounds validation

### Modifications:
- Added data validations
- Individual error handling per member
- New method getDistributionListMembersSafe()

**Date:** 2025-02-01  
**Author:** Jean Traullé