# SharePoint List Setup Instructions

## Changes Made to the Code

I've updated the `handleSubmit` function to:
1. Check if the list exists before trying to add items
2. Retrieve and display all available fields in the console for debugging
3. Dynamically map field names by trying common variations
4. Provide better error messages to help identify issues
5. Fixed TypeScript compilation errors

## SharePoint List Requirements

You need to create a SharePoint list named **"cloudlist"** with the following columns:

### Required Columns:
1. **Title** (single line of text) - System field, automatically created
2. **CustomerName** or **customer_name** (single line of text)
3. **Age** or **age** (Number field)
4. **Email** or **email** (single line of text)
5. **Rate** or **rate** or **Rating** (Number field - for rating 1-5)
6. **Comments** or **comments** (multiple lines of text) - Optional field

### Steps to Create the List:

1. Go to your SharePoint site
2. Click **Settings** (gear icon) → **Site Contents**
3. Click **New** → **List** or **List** (depending on your version)
4. Name the list: **cloudlist**
5. Add the columns:
   - **CustomerName** - Single line of text
   - **Age** - Number
   - **Email** - Single line of text
   - **Rate** - Number
   - **Comments** - Multiple lines of text (optional)

### Important Notes:

- The code will log all available fields to the browser console to help you identify the correct internal names
- If your fields have different internal names, the console will show them
- Make sure the user has permissions to add items to the list
- Field names are case-sensitive in SharePoint

## Testing the Solution

1. Open your browser's Developer Console (F12)
2. Fill out and submit the form
3. Check the console for:
   - List confirmation message
   - Available fields list
   - The data being sent
   - Any error messages

## Common Issues and Solutions:

### Issue: "List not found"
- Make sure the list name is exactly "cloudlist" (case-sensitive)
- Verify the list exists in the same site where the web part is deployed

### Issue: "Permission denied"
- Ensure the current user has "Contribute" or "Edit" permissions on the list
- Check list settings → Advanced settings → Item-level permissions

### Issue: "Field doesn't exist"
- Check the console log for the actual internal field names
- Update the code with the correct field names if they differ

### Issue: Field names don't match
- The code tries several variations automatically
- Check console logs to see which fields are available
- If needed, update the field name checks in the code

## Console Debugging

The form now logs detailed information to the browser console:
- List title confirmation
- All available fields with their internal names
- The data being sent to SharePoint
- Detailed error messages if submission fails

Use this information to identify any issues with your list configuration.

