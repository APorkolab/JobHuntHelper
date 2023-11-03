# JobHuntHelper
This repository contains a Google Apps Script designed to help job seekers manage and track their job applications within a Google Spreadsheet. The script automates follow-ups, manages application statuses, and organizes job search efforts with customizable settings. For a detailed explanation and walkthrough, check out the associated Medium blog post at [my Medium blog](https://slumbersolver.medium.com/job-search-juggler-automate-your-application-follow-ups-bfbf6631a39a).

## Features

- **Automated Follow-Ups:** Set reminders for follow-up emails after a specified number of days.
- **Application Status Tracking:** Easily update and view the status of each application.
- **Interview Scheduling:** Keep track of upcoming telephone and in-person interviews.
- **Documentation Management:** Record what documents were attached to each application.
- **Salary Expectation Handling:** Note down if salary expectations were requested and what was provided.

## How to Use

1. **Set Up the Spreadsheet:**
   - Create a new Google Spreadsheet.
   - Define columns as per the job application tracking needs (e.g., Company Name, Position, Application Date, etc.).

2. **Install the Script:**
   - Access the script editor from the spreadsheet by clicking on `Extensions > Apps Script`.
   - Copy and paste the provided Google Apps Script into the editor.
   - Save the script with a suitable name.

3. **Run the Script:**
   - Run the `createTriggers` function within the script editor to set up the necessary triggers.
   - Allow any permissions requested by the script to interact with your Google Spreadsheet and other services.

4. **Input Application Details:**
   - Enter the details of each job application into the respective columns.
   - Use the provided format to ensure the script functions correctly.

5. **Monitor Your Applications:**
   - The script will automatically change the colors of rows based on the application status.
   - Follow-up emails can be drafted automatically after a set number of days.

6. **Update Status Manually:**
   - Manually update the status of your applications as you receive responses or attend interviews.

## Columns Description

- `Company Name`: The name of the company where the application was sent.
- `Position`: The title of the position applied for.
- `Application Form`: How the application was submitted (e.g., online form, email).
- `Language`: The language of the application or the interview.
- `Application Time`: The date when the application was submitted.
- `Warning Message Column`: A reminder or alert set for yourself to follow up if you haven't heard back within a certain time frame.
- `Job Description PDF Link`: A hyperlink to the PDF containing the job description, which can be a useful reference for interview preparation.
- `Number of Interested Messages`: The count of communications you've received from the company, indicating their interest or responses to your application.
- `Name of the Contact Person`: The individual at the company who is your main point of contact, usually in HR or the hiring manager.
- `Contact E-mail Address`: The email address of your contact person for follow-ups or additional communication.
- `Job Link`: A direct link to the job posting on the company's career page or job board.
- `Comment`: Personal notes or initial impressions about the job or the application process.
- `What I Attached`: Documentation or files you've included with your application, such as a resume, cover letter, portfolio, etc.
- `They Asked for a Wage Claim`: Indicates whether the company has requested your salary expectations.
- `What I Gave`: Your response to the salary expectation inquiry, if applicable.
- `Feedback Time`: The date when you received feedback on your application, if any.
- `Telephone Interview`: The date and time of a scheduled telephone interview.
- `1st Interview`: The date and time of your first in-person or video interview.
- `2nd Interview`: The date and time of your second interview, if applicable.
- `3rd Interview`: The date and time of your third interview, if applicable.
- `Test/Trial Assignment`: Details about any practical tests or assignments given to you as part of the application process.
- `Offer`: Information on the job offer received, including the date of the offer and any pertinent details.

## Installation

1. Open your Google Spreadsheet where you wish to track job applications.
2. Click on `Extensions > Apps Script` to open the script editor.
3. Copy the script from this repository and paste it into the script editor.
4. Save and name your project in the script editor.
5. Run the `createTriggers` function to set up the automated features.

## Associated Medium Blog Post

For a comprehensive guide and best practices on using this job application tracker, please refer to our Medium blog post: [Job Search Juggler: Automate Your Application Follow-Ups](https://link.medium.com/DvmGNw6ZpEb).

## Contributing

Contributions to this project are welcome. Please fork the repository and submit a pull request with your updates.

## License

This project is open-sourced under the MIT License. See the [LICENSE](LICENSE) file for more details.

## Contact

For support or queries, please open an issue in the repository, and I will get back to you as soon as possible.

Happy job hunting!
