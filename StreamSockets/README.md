### Stream Sockets SSL Demo (here we are connecting to the Gmail SMTP server and issue various SMTP commands as an example)

<img width="578" height="562" alt="StreamSockets" src="https://github.com/user-attachments/assets/3de2491a-18ee-45da-abca-d08e7a3c44b3" />

The screenshot above shows how a usual session might unfold:

- We start with the ubiquitous HELO greeting (or its dyslexic sibling EHLO).
- Then we attempt authentication with AUTH LOGIN to which Gmail says no, we need an encrypted connection for that so we comply with the STARTTLS command!
- After that the AUTH LOGIN command works fine and we are prompted for a username and password (these need to be BASE64 encoded).
-Here we enter bogus credentials but a valid username and app password (not your main Gmail password!) will work correctly!
- Next steps would be sending various MIME encoded headers composing your actual message and recipients list (not included here for brevity).
- Finish with QUIT!
