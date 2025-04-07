Please validate the Excel (.xlsx) document attached and share your approval. Any necessary adjustments should be responded to in this email, and after the adjustments, a new document will be generated for validation and approval.

### How to validate users, groups, and access in the organization's technology environment:

1. **Users Tab**: Check if all users should exist in the environment based on their email (column A) or username (column K). Users who have been terminated can only be present in the list if column "E" is "Disabled", indicating that the user is blocked. Note that users who do not belong to a natural person (such as generic/non-nominal users) should not be maintained.

2. **SharePointMemberships Tab**: Check if all SharePoints should exist in the environment based on their name (column A). Each user with access to the SharePoint is listed in a new row; verify if the listed user (column D, E, F) should have access to the SharePoint indicated in the same row (column A). Also, check if each user's permission is appropriate for their role (column G).

3. **DLMemberships Tab**: Check if all Distributiol lists should exist in the environment based on their name (column A). Each user that receive messages from each Distribution List is listed in a new row; verify if the listed user (column D, E, F) should get e-mail messages from related Distribution List indicated in the same row (column A). Also, check if each user's permission is appropriate for their role (column G).

   **Permission Descriptions**:
   - **Owner**: Admin permission, has full control over the SharePoint/Distribution List.
   - **Member**: Have full control over files/get messages but cannot approve, add or remove new members
