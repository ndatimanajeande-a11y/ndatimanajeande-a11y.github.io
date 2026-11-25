# Ripple Rwanda customer management setup

## What this adds
- A `Customers` tab with one customer profile per normalized phone number.
- A `Bookings` tab that keeps every booking, including the customer ID, phone, source and status.
- An `Activity` tab that records when a booking is added.
- Website bookings are labelled `Website`.
- Staff can enter a WhatsApp enquiry through a PIN-protected Apps Script portal; those rows are labelled `WhatsApp`.
- Rwanda phone formats such as `078...` and `+25078...` are normalized so repeat customers can be matched.

## Required one-time Google Apps Script update
The website's current endpoint URL is already configured in `index.html`. The code in `Code.gs` must be copied into the same Apps Script project that owns that endpoint.

1. Open the existing Google Apps Script project connected to the Ripple Bookings spreadsheet.
2. Open `Code.gs`, replace its contents with this repository's `ripple-rwanda/Code.gs`, and save.
3. In **Project Settings > Script properties**, add:
   - Property: `RIPPLE_ADMIN_PIN`
   - Value: a strong private PIN known only to Ripple staff.
4. Click **Deploy > Manage deployments**.
5. Edit the existing web app deployment, select **New version**, and deploy. Keep the same deployment URL so the website continues posting to it.
6. Open the staff portal link in the website footer and test one WhatsApp intake using a test customer record. Check the spreadsheet's `Customers`, `Bookings`, and `Activity` tabs.

## How to record a WhatsApp customer
1. Open the Ripple website and select **Staff WhatsApp Intake** in the footer.
2. Enter the staff PIN.
3. Copy the customer's name, phone, pickup, destination, parcel size and requested time from their WhatsApp message.
4. Submit. The portal returns a booking reference and customer ID.
5. If the same customer contacts Ripple again using the same phone number, the system updates their existing customer profile and adds another booking row.

## Important operating notes
- The public website cannot automatically read WhatsApp chats. Automatic chat import requires a WhatsApp Business Platform/API integration and an approved webhook.
- The staff intake portal is protected by a server-side Script Property PIN. Do not put the PIN in website HTML or share it with customers.
- The website and tracking interface remain a prototype. A saved request is not proof that a rider is assigned or a parcel is delivered. Staff must confirm availability and update status.
- Collect only information needed to arrange the parcel. Limit spreadsheet access to authorized staff and do not store payment credentials or unrelated sensitive details.
- USSD is not active. Add USSD as a source only after an official shortcode/provider and its integration are configured.
