Renewal Review Builder
No-admin portable app

How to start
1. Unzip the RenewalReviewBuilderPortable folder.
2. Open the folder.
3. Double-click Start Renewal Review Builder.cmd.
   You can also double-click RenewalReviewBuilder.App.exe directly.

How to create a review pack
1. Select the Manulife renewal PDF.
2. Select the Manulife XLSM workbook.
3. Select the Westland PowerPoint template, if available.
4. Choose an output folder.
5. Click Validate.
6. Review the findings.
7. Click Export Pack.

Outputs
- Renewal_Review.xlsx
- Client_Renewal_Draft.pptx

PowerPoint premium overview table placeholders
- Put <<RateTable_Master>> or <<PremiumOverviewTable>> in a text box where the combined premium overview table should go.
- Put <<PremiumOverview2>> or <<PremiumOverview2Table>> where the combined premium overview should include the Single/Couple/Family health and dental breakdown.
- Put <<HealthDentalBreakdown>> if you only want the health and dental tier breakdown as a separate smaller table.
- Put <<RateTable_Divisions>> where all account/class divisions should be listed separately.
- Put <<RateTable_Division1>>, <<RateTable_Division2>>, and so on if you want one division per slide.
- The text box should contain only the placeholder text, with no other words in the same box.
- The placeholder text box controls where the table appears and how much space it has.

Notes
- No admin rights are required.
- Nothing is installed into Program Files.
- The app stores settings and logs here:
  %LOCALAPPDATA%\Westland\RenewalReviewBuilder
- If the PowerPoint template is locked or unavailable, the app still exports the Excel review workbook and creates a fallback draft deck.

If Windows blocks the app
1. Move the folder to Desktop or Documents and try again.
2. Right-click RenewalReviewBuilder.App.exe, choose Properties, and click Unblock if shown.
3. If company policy blocks unsigned apps, ask IT to allow the folder/app.
