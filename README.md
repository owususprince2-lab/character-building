# Character Building Programme LMS — Prototype

A professional blue/white front-end prototype for the Character Building Programme.

## Demo accounts
- Instructor: `bkoomson` / `BK123456`
- Student: register through the demo flow; the prototype creates a username from the student's name and a generated password.

## Important production work still required
This prototype intentionally does not process real money, store passwords securely, or send real email. Before public launch, connect:
- PostgreSQL/Supabase database
- Secure authentication and password hashing
- Verified payment gateway for Ghana card + Mobile Money
- Payment webhooks and receipts
- Cloud file storage
- Real email service
- Role-based authorization
- Assignment uploads, grading and feedback
- Push/in-app notifications
- Admin portal and audit logs
- Privacy, terms, backups and security controls

The 12 module names are placeholders because the exact module titles were not supplied in this conversation. Replace them in `app.js` in the `modules` array.
