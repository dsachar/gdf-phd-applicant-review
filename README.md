# PhD Applicant Evaluation Dashboard

A local, client-side web application designed to help reviewers efficiently evaluate, score, and rank PhD candidates for the GreenFieldData project.

## How to Run the App

Because this is a static, local dashboard, no installation or active internet server is required. All data processing and storage happens safely and privately within your own browser.

Simply double-click the **`index.html`** file in this folder, and it will open directly in your default web browser (Chrome, Firefox, Safari, Edge, etc.).

---

## 1. Exporting Your Data from EUSurvey

Before using the dashboard, you need to download the raw applicant data and their attached files (CVs, cover letters, transcripts) directly from the survey platform.

You can do this by exporting your data as an Excel file and choosing to include uploaded attachments. Here is how to perform the export:

1. **Access your survey:** Log in to [EUSurvey](https://ec.europa.eu/eusurvey/) and open the specific survey you are managing.
2. **Go to Results:** Click on the **Results** tab in the survey's top menu.
3. **Initiate Export:** Click the **Export** button.
4. **Select Format:** Choose **Excel (.xlsx)** as your desired data export format. 
5. **Include Attachments:** Ensure you check the option to include **Uploaded files**. EUSurvey will package the spreadsheet and all respondent-uploaded files together into a single **.zip** archive.
6. **Download:** Confirm the export. Once the system finishes processing, download the resulting archive to your computer. 

Extracting the downloaded `.zip` file will give you access to the `.xlsx` spreadsheet containing the structured survey answers, alongside a folder containing all the related files uploaded by the participants.

---

## 2. Loading Data into the Dashboard

1. Open `index.html` in your browser.
2. Click the **Select Survey Folder** button and select the **entire folder** you extracted from the EUSurvey download.
3. The dashboard will automatically find the spreadsheet, identify all documents, and load the candidates.
4. *Note: If multiple spreadsheets are found, you will be prompted to pick the correct one before starting.*

---

## 3. Using the Dashboard

- **Select an Applicant:** Use the left sidebar to search, sort, or browse through candidates. Click on a candidate to view their complete profile on the right.
- **Review Profile:** The right panel displays the applicant's educational background, Master's thesis details, reference emails, language proficiencies, and color-coded technical skills. Any external URLs provided for thesis downloads are also easily accessible.
- **Evaluate:** In the top right corner, use the "Rating & Evaluation" panel to assign integer scores (0-10) for BSc Grade, MSc Grade, Research Experience, and Fit to PhD. The total score will calculate automatically based on the assigned weights.
- **Adjust Weights:** Click **Adjust Weights** in the left sidebar to change how heavily each of the grading categories impacts the final Total Score.
- **Track Progress:** Check the **Hide Evaluated** box in the sidebar to temporarily remove applicants you've already scored from the list. You can also sort the list by Name, Rating, or Original Order.
- **Export Ratings:** Once you have finished evaluating, click **Export Ratings** in the sidebar to download a secure backup of your scores to your computer.
