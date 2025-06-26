// Requires: node-fetch (npm install node-fetch@2)
// Usage: node freshdesk-remove-deleted-contacts.js

const fetch = require('node-fetch'); // For Node.js. In browser, use global fetch.

const apiKey = "4KgveJYUOWZ5mio5lhR";
const freshdeskDomain = "itsupport-journebrands.freshdesk.com";
const encodedAuth = Buffer.from(`${apiKey}:X`).toString('base64');
const headers = {
  Authorization: `Basic ${encodedAuth}`,
  'Content-Type': 'application/json'
};

// Retrieve all deleted contacts with pagination
async function getDeletedContacts() {
  let contacts = [];
  let page = 1;
  let keepGoing = true;

  while (keepGoing) {
    const url = `https://${freshdeskDomain}/api/v2/contacts?state=deleted&page=${page}&per_page=100`;
    try {
      const response = await fetch(url, { method: 'GET', headers });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const data = await response.json();
      if (Array.isArray(data) && data.length > 0) {
        contacts = contacts.concat(data);
        if (data.length < 100) {
          keepGoing = false;
        } else {
          page++;
        }
      } else {
        keepGoing = false;
      }
    } catch (error) {
      console.warn(`Failed to retrieve contacts on page ${page}:`, error);
      break;
    }
  }
  return contacts;
}

// Hard delete a contact by ID
async function hardDeleteContactById(contactId) {
  const url = `https://${freshdeskDomain}/api/v2/contacts/${contactId}/hard_delete`;
  try {
    const response = await fetch(url, { method: 'DELETE', headers });
    if (response.ok) {
      console.log(`Permanently deleted contact ID ${contactId}`);
    } else {
      console.warn(`Failed to hard delete contact ID ${contactId}: HTTP ${response.status}`);
    }
  } catch (error) {
    console.warn(`Failed to hard delete contact ID ${contactId}:`, error);
  }
}

// Main execution
(async () => {
  const allDeletedContacts = await getDeletedContacts();
  console.log(`Total deleted contacts to permanently delete: ${allDeletedContacts.length}`);
  for (const contact of allDeletedContacts) {
    await hardDeleteContactById(contact.id);
  }
})();