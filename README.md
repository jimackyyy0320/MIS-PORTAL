# MIS-PORTAL: Student Records Portal

A web-based Academic Performance Tracker built using Google Apps Script (GAS) and Google Sheets. The portal allows students to view their grades, attendance, class schedules, and reviewers, while providing faculty with tools to manage rosters, edit grades, and generate attendance reports.

## Features

### 🎓 Student Portal
- **Dashboard & Profile:** View overall performance, edit personal profile data, and upload avatars.
- **Attendance Tracking:** Visual breakdown of attendance using Chart.js with calendar view.
- **Progress Tracking:** Daily progress charts (bar graphs) and major assessment scores (SA1, SA2, PE).
- **Reviewers Module:** Access and view subject reviewers directly from Google Drive. Can preview HTML files, generate PDFs, or download files.
- **Relative Evaluation:** See comparative academic performance (e.g., "Exceeds Expectations" based on Z-Score).
- **Smart Scheduling:** Real-time floating timer counting down to the next class, with an automated "In-Class" warning overlay that restricts phone usage during class hours.
- **Security features:** OTP-based password resets via email. Downloadable Student QR Code.

### 🏫 Faculty LMS
- **Master Roster:** View all enrolled students in a specific subject.
- **Inline Grade Editing:** Click on a student's score in the table to update it directly in the Google Sheet.
- **Student View:** Access an individual student's complete dashboard to monitor their progress.
- **Report Generation:** Generate daily attendance reports (Morning/Afternoon sessions) and export them as PNG images using `html2canvas`.

### ⚙️ System
- **Real-time Auto-Refresh:** Automatically polls the server every 60 seconds to keep data updated.
- **Dynamic Subject Configuration:** Subjects and their corresponding Google Sheets are managed centrally in a Master Sheet.
- **Marquee Announcements:** Banner at the top of the portal for URGENT or CASUAL announcements.
- **Maintenance Mode:** Retro TV "Offline" screen when the system status is set to MAINTENANCE or CLOSED.

---

## Known Bugs & Technical Limitations

1. **Insecure Authentication State:** Session state is managed purely on the frontend via `sessionStorage` (`{"id": "...", "isFaculty": false}`). A malicious user could manually edit `sessionStorage` in the browser console to view another student's data since the GAS backend functions (e.g., `fetchSubjectData`) do not enforce a secure server-side session token.
2. **Plain-text Passwords:** Passwords are saved in the "Credentials" Google Sheet in plain text, which poses a severe security risk.
3. **Google Apps Script Timeouts:** Heavy usage of `SpreadsheetApp.openById()` inside loops and processing large arrays of data (e.g., generating attendance for all subjects) can lead to the 6-minute execution time limit if the user base grows.
4. **Concurrency Issues:** Multiple students registering simultaneously or faculty updating grades at the exact same time might lead to Google Sheets race conditions or overwrite locks.
5. **Monolithic Architecture:** `Index.html` contains over 3,000 lines of mixed HTML, CSS, and JS. This monolithic structure makes it very difficult to debug, test, and maintain.
6. **Drive Searching Bottlenecks:** Searching Google Drive for avatars (`title contains "ID"`) on login is slow. As the number of images increases, this search will degrade performance significantly.

---

## Proposed Improvements

### 🎨 UI & Code Maintainability
- **Split Codebase:** Separate the monolithic `Index.html` into `Index.html`, `CSS.html`, and `JS.html`. Use GAS templating (`HtmlService.createTemplateFromFile`) to include them. This will make the frontend code much easier to manage.
- **State Management:** Implement a lightweight frontend framework (like Alpine.js or Vue.js) to manage the UI state. The current manual DOM manipulation (`document.getElementById().innerText = ...`) is brittle and prone to bugs as the app scales.

### 🛡️ Security
- **Password Hashing:** Implement a hashing algorithm (e.g., SHA-256) on the backend so passwords are never stored in plain text.
- **JWT / Server-side Sessions:** Generate a secure token upon login, pass it to the frontend, and verify it on every subsequent `google.script.run` call to ensure users can only fetch their own data.

### 🚀 Stability & Scalability
- **Database Migration:** If the system is expected to handle more than a few hundred students concurrently, migrate the database from Google Sheets to **Firebase Realtime Database** or **Cloud Firestore**. They provide native GAS integrations, handle real-time concurrency flawlessly, and eliminate read/write bottlenecks.
- **Store File URLs:** Instead of searching Google Drive for the avatar ID every time a user logs in, store the Avatar URL directly in a dedicated column in the `Credentials` sheet upon upload.
- **Batch Operations:** For faculty views, minimize Google Sheets API calls by utilizing `getValues()` and `setValues()` on entire ranges at once, rather than updating cell by cell, which is slow.
- **Rate Limiting:** The 60-second auto-refresh (`startAutoRefresh`) can easily exhaust Google's daily quota for `google.script.run` calls if many students keep the portal open. Consider increasing the interval or using WebSockets/Firebase listeners if real-time data is critical.
