// Function to make specific labels bold (e.g., "Start Time:") in a Google Doc body
function boldLabel(body, labelText) {
  const paragraphs = body.getParagraphs(); // Get all paragraphs in the document
  for (let i = 0; i < paragraphs.length; i++) {
    const textElement = paragraphs[i].editAsText(); // Convert paragraph to editable text
    const text = textElement.getText(); // Get the paragraph text
    const index = text.indexOf(labelText); // Find the label text
    if (index !== -1) {
      // If label found, make that portion bold
      textElement.setBold(index, index + labelText.length - 1, true);
    }
  }
}

// Triggered automatically when a Google Form is submitted
function onFormSubmit(e) {
  const values = e.values; // Get all submitted values from the form

  // Extract each value from the form response array using correct indices
  const date = values[1];
  const rawShiftTime = values[2];
  const company = values[3];
  const jobNumber = values[4];
  const rawStartTime = values[5];
  const startLocation = values[6];
  const instructions = values[7];
  const notes = values[8];
  const truckNumbersRaw = values[9]; // e.g. "RT03, RT12" 
  const tolls = values[10];
  const add01 = values[11];
  const add02 = values[12];
  const startTime02Raw = values[13];
  const startLocation02 = values[14];
  const instructions02 = values[15];
  
  // Convert comma-separated string to array
  const truckNumbers = truckNumbersRaw.split(',').map(truck => truck.trim()); // Multiple truck numbers 

  Logger.log(values);  // Add this line right after: const values = e.values;

  // Logging for debugging (viewable in script editor's logs)
  console.log("Raw Shift Time:", rawShiftTime);
  console.log("Raw Start Time:", rawStartTime);
  console.log("Raw Start Time 02:", startTime02Raw);

truckNumbers.forEach(truckNumber => { // Loop for each truck number 

  // At the top of onFormSubmit
  const formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Format times to be more readable
  const shiftTime = formatTime(rawShiftTime);
  const startTime = formatTime(rawStartTime);
  const startTime02 = formatTime(startTime02Raw);

  const sortableStartTime = convertToSortableTime(rawStartTime); // Defining sortableStartTime - NEW

  function convertToSortableTime(rawTime) { // Helper Function - NEW
    if (!rawTime) return "0000";
    const match = rawTime.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
    if (!match) return "0000";

    let hour = parseInt(match[1], 10);
    const minutes = match[2];
    const period = match[3] ? match[3].toUpperCase() : "";

    if (period === "PM" && hour < 12) hour += 12;
    if (period === "AM" && hour === 12) hour = 0;

    return `${hour.toString().padStart(2, '0')}${minutes}`; // End Helper Function - NEW
  }

  // ID of the Google Docs template used for archive copies
  const templateId = "1Gx6VGwAt202ffjtdHUJSF1ZCGlfS5ZluLFor_bp1Ggs";

  // Fallback archive folder ID if truck number doesn’t match
  const archiveFolderId = "1_idlRE8_vOcQSEX97k9iS1huoWf-ZbRb";

  // Map truck numbers to their specific live Google Doc IDs
  const truckDocIds = {
    "DT02": "1MiMefcafZbpwXz7K9uNlUXESZOx2GMT42Hev02gpKdM",
    "RT03": "1M4eJww7DaLG25CVKlt0uaRulLOFsWflH3e9z6FA9O5Q",
    "RT12": "1MhmPdyDzTu1NcXznoKRfgi1yMPHXMeryZGEww7QSH5k",
    "RT33": "1v1JM90NrT9oH7h55wfxigZjdys4yk0RcNc-0qlTGb78",
    "EXCL1": "1swlDuc93YWMZ4llSDNc2rMJi9-I4PhwRST_C7Fq3i2c",
    "WAJA01": "15bx7Ash8LlyuVQhNOdnbjifi6lC9OXVzYe7s_RsxYl8",
    "WAJA03": "1jLIpVCwaCumqimROSh4yhEjAp_3C7beC-vfU7QmHVPc",
    "IDR25": "1DwOvot33hjEZIMMTpU-MEq496oOgSEiGSx8wD08ltyg",
    "MANA1": "1oxWaEUwHCubVuFN-ET0mCuUTg9tDfceZknRfCb8tfuk",
    "MANA4": "1wPQACFX5w74ifwflf8AiSNUTsBZ6oAnMn1HncXeint0",
    "CHINOS1": "1od43UH5ATks90xNvJpEPiiqwWELBsQUf4Mrr7YVyM3s",
    "CHINOS5": "1TWgttIBvru6WGhzU-FgkxAEIy9YGfMesrygqA6lgn1M",
    "CHINITOS1": "1lZ-aRPEOUM5vpgxqDr7Iz42yT6jDk6C_2Df82qOd9nE",
    "CHINITOS5": "1jSAzQwIPo1TVKSJfAw44aUMdhQHJJIhoy32tddXXW3Y",
    "CHINITOS6": "1oysz6yP9ftBzz09843a3YU89pb7DO0uti7dY88zQhMo",
    "CONIG22": "1LdhX8cz-MVXp-ujFIUZ4wEKjoMoBq8rU16ujaAc3U4o",
    "CONIG29": "1kzKx6EkMjLPmxMGgzxLoyk88j9X6Sn-OdbrVL39x71Q",
    "CRAVO27": "10IbtW-6rhnbHldQFKnjI-GK3w57BrpOWLvkbUiCcgmo",
    "ECE03": "1HQgNe_XG7uLjaptVQB-VCM92D7Q4BaU6NM-bSmOjsQU",
    "RMT05": "17LVPMKZHgCc0MwQJpRrQFZBKaqb3_f88Tpk4A23BPC8",
    "RMT07": "1bkYeWuuRomaU7BiQMJ7zE9lfsTgSZMNLlAOnTs-Ngj4",
    "RMT77": "1L9xBt54pGobuCV-p7rsh943eR4ONxt3QI3UOyAXqNlw",
    "GEMA1": "1bFJMlS-6gOG0_-E8pf-ALPWHLVKN_f2SC693ppriVgU",
    "GEMA4": "18rzJfZYpQ0wg9OmtLnWT3GfgSC7siVhm-rwHfAOMz9M",
    "WAC01": "1gsTgt0cIpyFy-owPe4sGV9sQx4a8ZSSwRmUNwYFnH7s",
    "WAC02": "1g7XaB74sNAIF6s8YUkMuf4KtsXpKZ23aNys5HTQ4nN4",
    "WAC03": "1GKqObdhwbiqmHLkSdlrKbWo9o2VaiOKnbJig_N3QJ5Q",
    "SFB98": "13x0ZKR3u5hkEsRuayc5PP63PgASiDlgW0oVZyIpNaEY",
    "DARBY96": "1txAxpx7B2hKMWgvvjr-IwkofJzJFVT4HtlwhS-uZaJ8",
    "DARBY95": "1nkGAJ93MghUxQzR_nRb_zs6tlL7snqB74Oq0UA9R6PQ",
    "SFB97": "1wZgbB52doHeKvrsuThS5Yk_d5zO9gCYAnS0FLwSj-aQ",
    "RKO16": "1H1ql3y75ED_ncGCMOaaNuEATqBxMgWPxnimT2335x48",
    "ELS05": "1g5OASsDLyHJRCjQeeBQVOEG--DZLlxgcuI-xWsVVohQ",
    "WMBA11": "12tlrX9qDoNHedx7fzvgevJbz19baPzUPjK7JMPebosU",
    "JJM01": "1wnpQcPVrHxKq7Db1R57fxfXoQ5UEHdf7CrHt0Qc0Zcw"
  };

  // Map truck numbers to their specific archive folders
  const truckArchiveFolders = {
    "DT02": "13py6BfAZPxkPD9K0KPdXkXufBccxPalC",
    "RT03": "1qNay0z8In_iALbomZib2EpzD4JC8uBiF",
    "RT12": "1O8D1wd9hcDKq5cDC4uI7KDNGl5jI4WgU",
    "RT33": "1-Cm6ZO70dNKqZH8TpREoLaUBSmGNSUAy",
    "EXCL1": "1cQVmSDVmkOXMJJRu-eMm4l3NrpPyG3ZB",
    "WAJA01": "1BuvktOu9_O1xD2gkigODru6pzMw8iba4",
    "WAJA03": "17M26Ils8DUF0VwG01CSfP8aqjttr9Qs8",
    "IDR25": "1RnxhKV8DrYQc6BsjkJhseN150AeumB2e",
    "MANA1": "1FtIDPkClT9hO-z0H3P0CNbaRlvk98OuI",
    "MANA4": "1iCz5in0G0bQWTiyOk2AI-UAanSfVhBtG",
    "CHINOS1": "1KihAts27rXH1tOrzXhNITl0WFa_FwvY1",
    "CHINOS5": "1NcD_-X8fRJtUmf9e3xYsJUGxF0uADLJu",
    "CHINITOS1": "1LivVaVSUMC5Zt95W1glvDmOtU3JOwGg5",
    "CHINITOS5": "1upNwOfboSkghhyoEFM0hMnCQ0PAiOyBX",
    "CHINITOS6": "1GTJiTo9s8e-tc6f_fIUrblpc5lcSHOAD",
    "CONIG22": "1XFVQUBj4XaAFsukLCY-xBQEhRse8dPqw",
    "CONIG29": "16bD8d_Wmddkrd9V4vckuLxkdZBaE9U9d",
    "CRAVO27": "1ZhlD1VhfI5kLw_B7_F7Dtw-Hr7_gJgs5",
    "ECE03": "1khQ3KIRAziqbN5TvYqFaL09YAPwGZrBl",
    "RMT05": "1AfJ1RWXj5PDdf0xc8VAKk_y1d6OBkiB7",
    "RMT07": "1sl1rVtKO69ytMWbVjztEuEFBcIV13Rws",
    "RMT77": "11un1EjYTKWMeM5v8ic4ScAdgFZhLlLsb",
    "GEMA1": "1FXlTYmYifI90cf4ppIoqOn4MRbFRaovS",
    "GEMA4": "1kw58OjxzFF3FkU5J71kgGEDcnRPDB7LC",
    "WAC01": "1-atlR7A-SX_QbM5eb9ZxjMiTXyUoCAs8",
    "WAC02": "1VPUqc94wri_w5PoQujIJEYXs5uCQlQL1",
    "WAC03": "1uBGGVySMr-z6jAnnNoqS--Rd-Q4jHvt2",
    "SFB98": "1a-zywmeWriYQqTtY52FfbEcty-3-iMBa",
    "DARBY96": "1iCq32M3Qsx39S7YiExQ-1ZBEbnFr0tN6",
    "SFB97": "1mnfAMGFS0QhrbtAo27-RndfwlEhxeBL3",
    "DARBY95": "1xCwCrRoYRec97pS6YSkkA3QSMN_SqwHM",
    "RKO16": "1YKJkaig2CHJtt5At9t7-YyE9LwkXwo6i",
    "ELS05": "1NuIVeyHT803zn1kml69WYxUrCfkNcSI-",
    "WMBA11": "1pIaZ8dW4UZD_xtM5q4zCA1N2Q0onT5VB",
    "JJM01": "1l4dujP4Kig5VtK4w-ssokzrv6rhmf22a"
    };  

  // NO-OP: removed unnecessary DocumentApp.create to avoid creating a blank doc in My Drive.
  // We update existing live truck docs by ID and create archive copies from the template below.
  // Create new doc from template
  // const doc = DocumentApp.create(`${formattedDate}_${sortableStartTime}_Dispatch_${truckNumber}_${jobNumber}`);
  // const body = doc.getBody();

  // Template block with placeholders to be replaced
  let placeholdersBlock = `✅ CONFIRM DISPATCH ✅

{{DATE}}
{{SHIFT TIME}}
{{COMPANY}}
Job No. {{JOB NUMBER}}

{{ADD 01}}
▪️Start Time: {{START TIME}}
▪️Start Location:
{{START LOCATION}}
▪️Instructions:
{{INSTRUCTIONS}}
{{ADD 02}}
{{START TIME 02:}}
{{START LOCATION 02:}}
{{INSTRUCTIONS 02:}}
{{NOTES}}

➡️ Working for Capital Contracting Group

#️⃣ {{TRUCK NUMBER}} #️⃣

🎧CB Radios must be ON and working at ALL times🎧

🚫NO STOPS while truck is loaded with asphalt🚫

⚠️ Check in/out with either (A)Trucking Foreman (B)Plant/Scale House (C)Earle Personnel (D)GPS Tracker(if assigned), or (E)FleetWatcher APP.

⚠️ Dump body must be clean or else truck will be sent home.

🚨 NOTES 🚨 
1️⃣ Tolls authorized: {{TOLLS}}
2️⃣ Remove towing hitch from the dump truck.
3️⃣ Please ensure you have safety vest/hard hat. Please follow construction safety rules, no speeding on job site.
4️⃣ It is the responsibility of the hauler to communicate, verify and confirm check in/out times with Earle personnel.
5️⃣ Breakdowns, call outs and late arrivals must be reported immediately.
6️⃣ Do not leave job early!
7️⃣ Drivers must report all problems to ALEX immediately! → (732) 470-8667`;

  // Replace check-in/out section with custom MCRC message if job is MCRC
  if (jobNumber.toUpperCase() === 'MCRC') {
    placeholdersBlock = placeholdersBlock.replace(
      /Check in\/out[\s\S]*?📱 FleetWatcher APP\n\n/,
      `*5 LOADS MINIMUM* ⚠️ 

🛑Monmouth County Reclamation Center gate closes at 03:30 PM!

⚠️Save all Pure Soil & Monmouth County Reclamation Center tickets and give to CCG. Tickets are required for payment!

*NO TICKETS - NO PAYMENT*\n\n`
    );
  }

  // Replace check-in/out section with custom CMPM message if job is CMPM
  if (jobNumber.toUpperCase() === 'CMPM') {
    placeholdersBlock = placeholdersBlock.replace(
      /Check in\/out[\s\S]*?📱 FleetWatcher APP\n\n/,
      `*2 ROUNDS MINIMUM* ⚠️ 

⚠️Save all Pure Soil, Colony Materials, & Plumstead Materials tickets and give to CCG. Tickets are required for payment!

*NO TICKETS - NO PAYMENT*\n\n`
    );
  }

  // Get the Google Doc ID for the truck's live doc
  const docId = truckDocIds[truckNumber];
  if (!docId) {
    console.log(`Unknown truck number: ${truckNumber}`);
    return;
  }

  // === UPDATE THE LIVE TRUCK DOC ===
  const truckDoc = DocumentApp.openById(docId);
  const truckBody = truckDoc.getBody();
  truckBody.setText(placeholdersBlock); // Set template block text
  
  // Bold + underline specific MCRC lines if present
  if (jobNumber.toUpperCase() === 'MCRC') {
    const bodyText = truckBody.editAsText();
    const phrasesToStyle = ["5 LOADS MINIMUM", "NO TICKETS - NO PAYMENT"];
    phrasesToStyle.forEach(phrase => {
      const idx = bodyText.getText().indexOf(phrase);
      if (idx !== -1) {
        bodyText.setBold(idx, idx + phrase.length, true);
        bodyText.setUnderline(idx, idx + phrase.length, true);
      }
    });
  }

  // Bold + underline specific MCRC lines if present
  if (jobNumber.toUpperCase() === 'CMPM') {
    const bodyText = truckBody.editAsText();
    const phrasesToStyle = ["2 ROUNDS MINIMUM", "NO TICKETS - NO PAYMENT"];
    phrasesToStyle.forEach(phrase => {
      const idx = bodyText.getText().indexOf(phrase);
      if (idx !== -1) {
        bodyText.setBold(idx, idx + phrase.length, true);
        bodyText.setUnderline(idx, idx + phrase.length, true);
      }
    });
  }

  // Replace placeholders with actual values from form
  truckBody.replaceText("{{DATE}}", date);
  truckBody.replaceText("{{SHIFT TIME}}", shiftTime + " ");
  truckBody.replaceText("{{COMPANY}}", company);
  truckBody.replaceText("{{JOB NUMBER}}", jobNumber);
  truckBody.replaceText("{{START TIME}}", startTime);
  truckBody.replaceText("{{START LOCATION}}", startLocation);
  truckBody.replaceText("{{INSTRUCTIONS}}", instructions);
  truckBody.replaceText("{{NOTES}}", notes);
  truckBody.replaceText("{{TRUCK NUMBER}}", truckNumber);
  truckBody.replaceText("{{TOLLS}}", tolls);

  // Handle optional field ADD 01
  if (add01.trim()) {
    truckBody.replaceText("{{ADD 01}}", add01.trim());
  } else {
    const paragraph = findParagraphContaining(truckBody, "{{ADD 01}}");
    if (paragraph) paragraph.removeFromParent();
  }

  // Handle optional field ADD 02
  if (add02.trim()) {
    const para = findParagraphContaining(truckBody, "{{ADD 02}}");
    if (para) {
      const index = truckBody.getChildIndex(para);
      truckBody.insertParagraph(index, ""); // Insert blank line above assignment 2 if it exists
    }
    truckBody.replaceText("{{ADD 02}}", add02.trim());
  } else {
    const paragraph = findParagraphContaining(truckBody, "{{ADD 02}}");
    if (paragraph) paragraph.removeFromParent();
  }

  // Format and bold assignment 2 fields if present, otherwise remove
  if (startTime02Raw.trim()) {
    const label = "▪️Start Time 02:";
    const formatted = `${label} ${startTime02}`;
    truckBody.replaceText("{{START TIME 02:}}", formatted);
    boldLabel(truckBody, label);
  } else {
    const paragraph = findParagraphContaining(truckBody, "{{START TIME 02:}}");
    if (paragraph) paragraph.removeFromParent();
  }

  if (startLocation02.trim()) {
    const label = "▪️Start Location 02:";
    const formatted = `${label}\n${startLocation02}`;
    truckBody.replaceText("{{START LOCATION 02:}}", formatted);
    boldLabel(truckBody, label);
  } else {
    const paragraph = findParagraphContaining(truckBody, "{{START LOCATION 02:}}");
    if (paragraph) paragraph.removeFromParent();
  }

  if (instructions02.trim()) {
    const label = "▪️Instructions 02:";
    const formatted = `${label}\n${instructions02}`;
    truckBody.replaceText("{{INSTRUCTIONS 02:}}", formatted);
    boldLabel(truckBody, label);
  } else {
    const paragraph = findParagraphContaining(truckBody, "{{INSTRUCTIONS 02:}}");
    if (paragraph) paragraph.removeFromParent();
  }

  // Bold core labels in the dispatch
  boldLabel(truckBody, "CONFIRM DISPATCH");
  boldLabel(truckBody, "Start Time");
  boldLabel(truckBody, "Start Location");
  boldLabel(truckBody, "Instructions");
  boldLabel(truckBody, "Start Time 02:");
  boldLabel(truckBody, "Start Location 02:");
  boldLabel(truckBody, "Instructions 02:");

  truckDoc.saveAndClose();

  // === CREATE ARCHIVE COPY ===
  try {
    const template = DriveApp.getFileById(templateId); // Get base template file
    const folderId = truckArchiveFolders[truckNumber] || archiveFolderId; // Use truck-specific folder or fallback to default
    const archiveFolder = DriveApp.getFolderById(folderId); // Get the target folder

    // Create a file name using sortable date and time for easy chronological sorting
    const sortableDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // Use the sortableStartTime already calculated from rawStartTime
    const newName = `${formattedDate}_${sortableStartTime}_Dispatch_${truckNumber}_${jobNumber}`; // Archive Name Format - NEW

    // Make a copy of the template into the archive folder and open the new document
    const archiveCopy = template.makeCopy(newName, archiveFolder);
    archiveCopy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // Access for anyone with link
    const archiveDoc = DocumentApp.openById(archiveCopy.getId());
    const archiveBody = archiveDoc.getBody();

    // Replace check-in/out section with custom MCRC message in archive if applicable
    if (jobNumber.toUpperCase() === 'MCRC') {
      const archiveText = archiveBody.getText();
      const updatedText = archiveText.replace(
         /Check in\/out[\s\S]*?📱 FleetWatcher APP\n\n/,
         `*5 LOADS MINIMUM*⚠️ 

🛑Monmouth County Reclamation Center gate closes at 03:30 PM!
 
⚠️Save all Pure Soil & Monmouth County Reclamation Center tickets and give to CCG. Tickets are required for payment!

*NO TICKETS - NO PAYMENT*\n\n`
      );
      archiveBody.setText(updatedText); 
    }

    // Replace check-in/out section with custom CMPM message in archive if applicable
    if (jobNumber.toUpperCase() === 'CMPM') {
      const archiveText = archiveBody.getText();
      const updatedText = archiveText.replace(
         /Check in\/out[\s\S]*?📱 FleetWatcher APP\n\n/,
         `*2 ROUNDS MINIMUM*⚠️ 
 
⚠️Save all Pure Soil, Colony Materials, & Plumstead Materials tickets and give to CCG. Tickets are required for payment!

*NO TICKETS - NO PAYMENT*\n\n`
      );
      archiveBody.setText(updatedText); 
    }

    // Now apply bold + underline styles
    const textObj = archiveBody.editAsText();
    const phrasesToStyle = ["5 LOADS MINIMUM", "2 ROUNDS MINIMUM", "NO TICKETS - NO PAYMENT"];
    phrasesToStyle.forEach(phrase => {
      const idx = textObj.getText().indexOf(phrase);
      if (idx !== -1) {
        textObj.setBold(idx, idx + phrase.length, true);
        textObj.setUnderline(idx, idx + phrase.length, true);
      }
    });
  
    // Replace basic placeholders
    archiveBody.replaceText("{{DATE}}", date);
    archiveBody.replaceText("{{SHIFT TIME}}", shiftTime + " ");
    archiveBody.replaceText("{{COMPANY}}", company);
    archiveBody.replaceText("{{JOB NUMBER}}", jobNumber);
    archiveBody.replaceText("{{START TIME}}", startTime);
    archiveBody.replaceText("{{START LOCATION}}", startLocation);
    archiveBody.replaceText("{{INSTRUCTIONS}}", instructions);
    archiveBody.replaceText("{{NOTES}}", notes);
    archiveBody.replaceText("{{TRUCK NUMBER}}", truckNumber);
    archiveBody.replaceText("{{TOLLS}}", tolls);

    // Handle optional ADD 01
    if (add01.trim()) {
      archiveBody.replaceText("{{ADD 01}}", add01.trim());
    } else {
      const paragraph = findParagraphContaining(archiveBody, "{{ADD 01}}");
      if (paragraph) paragraph.removeFromParent(); // Remove line if empty
    }

    // Handle optional ADD 02 (with blank line above if present)
    if (add02.trim()) {
      const para = findParagraphContaining(archiveBody, "{{ADD 02}}");
      if (para) {
        const index = archiveBody.getChildIndex(para);
        archiveBody.insertParagraph(index, ""); // Insert blank line above
      }
      archiveBody.replaceText("{{ADD 02}}", add02.trim());
    } else {
      const paragraph = findParagraphContaining(archiveBody, "{{ADD 02}}");
      if (paragraph) paragraph.removeFromParent();
    }

    // Handle START TIME 02 (bold label if present)
    if (startTime02Raw.trim()) {
      const label = "Start Time 02:";
      const formatted = `${label} ${startTime02}`;
      archiveBody.replaceText("{{START TIME 02:}}", formatted);
      boldLabel(archiveBody, label);
    } else {
      const paragraph = findParagraphContaining(archiveBody, "{{START TIME 02:}}");
      if (paragraph) paragraph.removeFromParent();
    }

    // Handle START LOCATION 02 (bold label if present)
    if (startLocation02.trim()) {
      const label = "Start Location 02:";
      const formatted = `${label}\n${startLocation02}`;
      archiveBody.replaceText("{{START LOCATION 02:}}", formatted);
      boldLabel(archiveBody, label);
    } else {
      const paragraph = findParagraphContaining(archiveBody, "{{START LOCATION 02:}}");
      if (paragraph) paragraph.removeFromParent();
    }

    // Handle INSTRUCTIONS 02 (bold label if present)
    if (instructions02.trim()) {
      const label = "Instructions 02:";
      const formatted = `${label}\n${instructions02}`;
      archiveBody.replaceText("{{INSTRUCTIONS 02:}}", formatted);
      boldLabel(archiveBody, label);
    } else {
      const paragraph = findParagraphContaining(archiveBody, "{{INSTRUCTIONS 02:}}");
      if (paragraph) paragraph.removeFromParent();
    }

    // Apply bold formatting to standard labels
    boldLabel(archiveBody, "CONFIRM DISPATCH");
    boldLabel(archiveBody, "Start Time");
    boldLabel(archiveBody, "Start Location");
    boldLabel(archiveBody, "Instructions");
    boldLabel(archiveBody, "Start Time 02:");
    boldLabel(archiveBody, "Start Location 02:");
    boldLabel(archiveBody, "Instructions 02:");

    archiveDoc.saveAndClose(); // Finalize and save the document
    console.log(`Archive copy created successfully: ${newName}`);
  } catch (error) {
    console.error("Error creating archive copy:", error); // Log any issues
  }

  // Format times to remove seconds
  function formatTime(rawTime) {
    Logger.log("Raw value received in formatTime: " + rawTime);
    if (!rawTime) return "";

    // Try to parse rawTime string like "5:00:00 AM"
    const timePattern = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i;
    const match = rawTime.match(timePattern);
    if (!match) return rawTime;  // Return raw if format unrecognized

    let [, hourStr, minuteStr, ampm] = match;
    let hour = parseInt(hourStr, 10);
    const minute = parseInt(minuteStr, 10);

    // Convert to 24-hour format if AM/PM exists
    if (ampm) {
      const isPM = ampm.toUpperCase() === "PM";
      if (isPM && hour !== 12) hour += 12;
      if (!isPM && hour === 12) hour = 0;
    }

    // Create a date with today's date, but custom hour and minute
    const date = new Date();
    date.setHours(hour);
    date.setMinutes(minute);
    date.setSeconds(0);
    date.setMilliseconds(0);

    return Utilities.formatDate(date, Session.getScriptTimeZone(), "h:mm a");
  }

  // Searches through all paragraphs in a Google Doc body to find the first paragraph
  // that contains a specific text string (e.g., a placeholder like {{ADD 01}}).
  // Useful for removing or inserting text conditionally during template replacement.
  // Returns the Paragraph object if found, or null if not found.
  function findParagraphContaining(body, searchText) {
    const paragraphs = body.getParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
      if (paragraphs[i].getText().includes(searchText)) {
        return paragraphs[i];
      }
    }
    return null;
  }

  // Retrieves and returns a sorted list of dispatch files for a given truck.
  // Dispatch files must follow this naming format:
  // "YYYY-MM-DD_HHMM_Dispatch_TRUCKNUMBER_JOBNUMBER"
  // Returns an array of objects: [{ date, time24hr, jobNumber, url }, ...]
  // Results are sorted chronologically for later display or filtering.
  function getTruckDispatches(truckNumber) {
    const truckFolders = {
      "DT02": "13py6BfAZPxkPD9K0KPdXkXufBccxPalC",
      "RT03": "1qNay0z8In_iALbomZib2EpzD4JC8uBiF",
      "RT12": "1O8D1wd9hcDKq5cDC4uI7KDNGl5jI4WgU",
      "RT33": "1-Cm6ZO70dNKqZH8TpREoLaUBSmGNSUAy",
      "EXCL1": "1cQVmSDVmkOXMJJRu-eMm4l3NrpPyG3ZB",
      "WAJA01": "1BuvktOu9_O1xD2gkigODru6pzMw8iba4",
      "WAJA03": "17M26Ils8DUF0VwG01CSfP8aqjttr9Qs8",
      "IDR25": "1TwREIe9wEBmq5EW3BFaA6d0LUlhu1n7W",
      "MANA1": "1FtIDPkClT9hO-z0H3P0CNbaRlvk98OuI",
      "MANA4": "1iCz5in0G0bQWTiyOk2AI-UAanSfVhBtG",
      "CHINOS1": "1KihAts27rXH1tOrzXhNITl0WFa_FwvY1",
      "CHINOS5": "1NcD_-X8fRJtUmf9e3xYsJUGxF0uADLJu",
      "CHINITOS1": "1LivVaVSUMC5Zt95W1glvDmOtU3JOwGg5",
      "CHINITOS5": "1upNwOfboSkghhyoEFM0hMnCQ0PAiOyBX",
      "CHINITOS6": "1GTJiTo9s8e-tc6f_fIUrblpc5lcSHOAD",
      "CONIG22": "1XFVQUBj4XaAFsukLCY-xBQEhRse8dPqw",
      "CONIG29": "16bD8d_Wmddkrd9V4vckuLxkdZBaE9U9d",
      "CRAVO27": "1ZhlD1VhfI5kLw_B7_F7Dtw-Hr7_gJgs5",
      "ECE03": "1khQ3KIRAziqbN5TvYqFaL09YAPwGZrBl",
      "RMT05": "1AfJ1RWXj5PDdf0xc8VAKk_y1d6OBkiB7",
      "RMT07": "1sl1rVtKO69ytMWbVjztEuEFBcIV13Rws",
      "RMT77": "11un1EjYTKWMeM5v8ic4ScAdgFZhLlLsb",
      "GEMA1": "1FXlTYmYifI90cf4ppIoqOn4MRbFRaovS",
      "GEMA4": "1kw58OjxzFF3FkU5J71kgGEDcnRPDB7LC",
      "WAC01": "1-atlR7A-SX_QbM5eb9ZxjMiTXyUoCAs8",
      "WAC02": "1VPUqc94wri_w5PoQujIJEYXs5uCQlQL1",
      "WAC03": "1uBGGVySMr-z6jAnnNoqS--Rd-Q4jHvt2",
      "SFB98": "1a-zywmeWriYQqTtY52FfbEcty-3-iMBa",
      "DARBY96": "1iCq32M3Qsx39S7YiExQ-1ZBEbnFr0tN6",
      "SFB97": "1mnfAMGFS0QhrbtAo27-RndfwlEhxeBL3",
      "DARBY95": "1xCwCrRoYRec97pS6YSkkA3QSMN_SqwHM",
      "RKO16": "1YKJkaig2CHJtt5At9t7-YyE9LwkXwo6i",
      "ELS05": "1NuIVeyHT803zn1kml69WYxUrCfkNcSI-",
      "WMBA11": "1pIaZ8dW4UZD_xtM5q4zCA1N2Q0onT5VB",
      "JJM01": "1l4dujP4Kig5VtK4w-ssokzrv6rhmf22a"
      };

    const folderId = truckFolders[truckNumber];
    if (!folderId) return [];

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const dispatches = [];

    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName(); // e.g. "2025-05-13_0615_Dispatch_RT03_00123"
      const url = file.getUrl();

      const match = name.match(/^(\d{4}-\d{2}-\d{2})_(\d{4})_Dispatch_[^_]+_(\d+)/);
      if (match) {
        dispatches.push({
          date: match[1],
          time24hr: match[2],
          jobNumber: match[3],
          url: url
        });
      }
    }

    // Sort by date and time
    dispatches.sort((a, b) => {
      const aKey = `${a.date}_${a.time24hr}`;
      const bKey = `${b.date}_${b.time24hr}`;
      return aKey.localeCompare(bKey);
    });

    return dispatches;
  }

  // Update each Truck to Company Folder
  const truckToCompanyMap = {
  'DT02': 'Capital Contracting Group',
  'RT03': 'RT Masonry',
  'RT12': 'RT Masonry',
  'RT33': 'RT Masonry',
  'EXCL1': 'Exclusive Contractors',
  'WAJA01': 'WAJA Contractors',
  'WAJA03': 'WAJA Contractors',
  'IDR25': 'IDR Trucking',
  'MANA1': 'MANA Concrete',
  'MANA4': 'MANA Concrete',
  'CHINOS1': 'Chinitos Trucking Service',
  'CHINOS5': 'Chinitos Trucking Service',
  'CHINITOS1': 'Chinitos Trucking Service',
  'CHINITOS5': 'Chinitos Trucking Service',
  'CHINITOS6': 'Chinitos Trucking Service',
  'CONIG22': 'Conigliaro Trucking',
  'CONIG29': 'Conigliaro Trucking',
  'CRAVO27': 'Cravo Enterprises',
  'ECE03': 'ECE Trucking',
  'RMT05': 'R Moreano Trucking',
  'RMT07': 'R Moreano Trucking',
  'RMT77': 'R Moreano Trucking',
  'GEMA1': 'RT Masonry',
  'GEMA4': 'RT Masonry',
  'WAC01': 'WA Construction',
  'WAC02': 'WA Construction',
  'WAC03': 'WA Construction',
  'SFB98': 'Yousef',
  'DARBY96': 'Yousef',
  'SFB97': 'Yousef',
  'DARBY95': 'Yousef',
  'RKO16': 'Yousef',
  'ELS05': 'Element Logistics',
  'WMBA11': 'WMBA Trucking',
  'JJM01': 'JJM Dump'
  // Add more as needed
};

const companyName = truckToCompanyMap[truckNumber];
if (companyName) {
  updateCompanyDispatchPage(companyName);
} else {
  Logger.log(`Company not found for truck number: ${truckNumber}`);
}

}); // End Loop for multiple trucks 
} // End onFormSubmit =============================================================================== End of onFormSubmit function

// Begin updateCompanyDispatchPage function on website
function updateCompanyDispatchPage(companyName) {
  const folderMap = {
    "Capital Contracting Group": ["DT02"],
    "RT Masonry": ["RT03", "RT12", "RT33", "GEMA1", "GEMA4"],
    "Exclusive Contractors": ["EXCL1"],
    "WAJA Contractors": ["WAJA01", "WAJA03"],
    "IDR Trucking": ["IDR25"],
    "MANA Concrete": ["MANA1", "MANA4"],
    "Chinitos Trucking Service": ["CHINOS1", "CHINOS5", "CHINITOS1", "CHINITOS5", "CHINITOS6"],
    "Conigliaro Trucking": ["CONIG22", "CONIG29"],
    "Cravo Enterprises": ["CRAVO27"],
    "ECE Trucking": ["ECE03"],
    "R Moreano Trucking": ["RMT05", "RMT07", "RMT77"],
    "WA Construction": ["WAC01", "WAC02", "WAC03"],
    "Yousef": ["SFB98", "DARBY96", "SFB97", "DARBY95", "RKO16"],
    "Element Logistics": ["ELS05"],
    "WMBA Trucking": ["WMBA11"],
    "JJM Dump": ["JJM01"]
  };

  const parentFolder = DriveApp.getFolderById("1_idlRE8_vOcQSEX97k9iS1huoWf-ZbRb");
  const companyFolder = parentFolder.getFoldersByName(companyName).next();
  const truckFolders = folderMap[companyName];

  let allFiles = [];

  // Aggregate dispatch files from all truck folders within the company folder
  for (const truckNum of truckFolders) {
    const truckFolder = companyFolder.getFoldersByName(truckNum).next();
    const files = truckFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().endsWith(".html")) continue; // skip HTML pages
      allFiles.push(file);
    }
  }

  // === Only show dispatches newer than 18 days, but don’t delete them ===
  const EIGHTEEN_DAYS_MS = 18 * 24 * 60 * 60 * 1000;
  const now = new Date().getTime();

  const recentFiles = [];
  for (const file of allFiles) {
    const createdTime = file.getDateCreated().getTime();
    if ((now - createdTime) <= EIGHTEEN_DAYS_MS) {
      recentFiles.push(file);
    }
  }
  allFiles = recentFiles; // use this filtered list for page generation


  // === Clean up files after 18 days ===
 // const EIGHTEEN_DAYS_MS = 18 * 24 * 60 * 60 * 1000;
  //const now = new Date().getTime();

  //for (const file of allFiles) {
    //const createdTime = file.getDateCreated().getTime();
    //if ((now - createdTime) > EIGHTEEN_DAYS_MS) {
      //try {
        //file.setTrashed(true); // Move to trash instead of deleting outright
        //Logger.log(`🗑️ Trashed old dispatch: ${file.getName()}`);
      //} catch (err) {
        //Logger.log(`❌ Error trashing file: ${file.getName()} - ${err}`);
      //}
    //}
  //}


  // Parse metadata and sort
  const dispatches = allFiles.map(file => {
    const name = file.getName();
    const dateMatch = name.match(/(\d{4}-\d{2}-\d{2})_(\d{4})/);
    const date = dateMatch ? new Date(`${dateMatch[1]}T${dateMatch[2].slice(0,2)}:${dateMatch[2].slice(2)}:00`) : new Date(0);
    return { file, name, date };
  }).sort((a, b) => a.date - b.date);

    function stripTime(date) {
      return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    }

    const today = stripTime(new Date());

  const upcoming = [], todayList = [], past = [];

  for (const d of dispatches) {
    const url = d.file.getUrl();
    const parts = d.name.split('_');
    const [dispatchDateStr, timeStr, , truckNumber, jobNumberWithExt] = parts;
    const jobNumber = jobNumberWithExt.replace(/\..+$/, '');

    const dispatchDate = new Date(`${dispatchDateStr}T${timeStr.slice(0, 2)}:${timeStr.slice(2)}:00`);
    const friendlyDate = dispatchDate.toLocaleDateString('en-US', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    });
    const friendlyTime = dispatchDate.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });

  let labelContent = `${friendlyDate} @ ${friendlyTime} – ${truckNumber} – ${jobNumber}`;
  let statusLabel = "";
  let labelStyle = "";

  if (d.name.includes("CANCELLED")) {
    statusLabel = " <span style='color: #d93025; font-weight: bold;'>CANCELLED</span>";
    labelStyle = "color: grey; text-decoration: line-through;";
  } else if (d.name.includes("AMEND")) {
    statusLabel = " <span style='color: #FFBF00; font-weight: bold;'>AMENDMENT</span>";
  }

  const label = `<span style="${labelStyle}">${labelContent}</span>${statusLabel}`;
  const link = `<div class="dispatch-block"><a href="${url}">${label}</a></div>`;

    const entry = { date: d.date, html: link };

    const dispatchDay = stripTime(dispatchDate);
    
    if (dispatchDay.getTime() < today.getTime()) {
      past.push(entry);
    } else if (dispatchDay.getTime() === today.getTime()) {
      todayList.push(entry);
    } else {
      upcoming.push(entry);
    }
  }
  upcoming.sort((a, b) => b.date - a.date);
  todayList.sort((a, b) => b.date - a.date);
  past.sort((a, b) => b.date - a.date);

  const upcomingHTML = upcoming.map(e => e.html).join('\n');
  const todayHTML = todayList.map(e => e.html).join('\n');
  const pastHTML = past.map(e => e.html).join('\n');


  //Webpage layout
  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>${companyName} Dispatches</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; background: #f9f9f9; }
    h1 {
      font-size: 28px;
      text-align: center;
      padding: 10px 20px;
      display: inline-block;
      background-color: #FFD700; /* construction yellow */
      border-radius: 25px;
      box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }
    .section {
      margin-top: 40px;
      padding: 20px;
      border-radius: 10px;
      background-color: #ffffff;
      box-shadow: 0 0 8px rgba(0,0,0,0.05);
    }
    .upcoming { border: 3px solid #4CAF50; } 
    .today { border: 3px solid #2196F3; }
    .past { border: 3px solid #9E9E9E; }
    h2 {
      font-size: 20px;
      margin-top: 0;
    }
    .dispatch-block {
      padding: 10px 0;
      border-bottom: 1px solid #e0e0e0;
    }
    .dispatch-block a {
      text-decoration: none;
      color: #333;
    }
    .dispatch-block a:hover {
      text-decoration: underline;
    }
    .dispatch-block a {
      color: #1a73e8; /* Google Blue */
      text-decoration: underline;
      font-weight: bold; /*Make links bold */
    }
    .title-container {
      text-align: center;
      margin-bottom: 24px;
    }
    .dispatch-title {
      display: inline-block;
      background-color: #FFD700; /* Construction yellow */
      padding: 12px 24px;
      border-radius: 999px; /* Pill/bubble shape */
      font-size: 2em;
      font-weight: bold;
      color: #000;
      box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    }
  </style>
</head>
<body>
    <div class="title-container">
      <h1 class="dispatch-title">${companyName} Dispatches</h1>
    </div>

  <div class="section upcoming">
    <h2>Upcoming</h2>
    ${upcomingHTML || '<p>No upcoming dispatches.</p>'}
  </div>

  <div class="section today">
    <h2>Today</h2>
    ${todayHTML || '<p>No dispatches for today.</p>'}
  </div>

  <div class="section past">
    <h2>Past</h2>
    ${pastHTML || '<p>No past dispatches.</p>'}
  </div>
</body>
</html>`;

  const fileName = `${companyName.replace(/\s+/g, '_')}_dispatch_list.html`;
  const existing = companyFolder.getFilesByName(fileName);
  let htmlFile;

  if (existing.hasNext()) {
    htmlFile = existing.next();
    htmlFile.setContent(html);
  } else {
    htmlFile = companyFolder.createFile(fileName, html, MimeType.HTML);
  }

  htmlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}


// Adds update button to google sheets for updating all pages manually
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚚 Dispatch Tools')
    .addItem('🔄 Refresh All Company Pages', 'updateAllCompanyPages')
    .addToUi();
}

 // updates all pages based on trigger -- every 1 hour
function updateAllCompanyPages() {
  const companies = [
    "Capital Contracting Group",
    "RT Masonry",
    "Exclusive Contractors",
    "WAJA Contractors",
    "IDR Trucking",
    "MANA Concrete",
    "Chinitos Trucking Service",
    "Conigliaro Trucking",
    "Cravo Enterprises",
    "ECE Trucking",
    "WA Construction",
    "Yousef",
    "R Moreano Trucking",
    "Element Logistics",
    "WMBA Trucking",
    "JJM Dump"
  ];

  for (const company of companies) {
    updateCompanyDispatchPage(company);
  }
}

function testClaspConnection() {
  Logger.log("Clasp connection working");
}
