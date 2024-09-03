/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
// Set the minimum date for the date accessed input to today
function setMaxDate() {
  const today = new Date();
  const dd = String(today.getDate()).padStart(2, "0");
  const mm = String(today.getMonth() + 1).padStart(2, "0"); // January is 0!
  const yyyy = today.getFullYear();
  const maxDate = `${yyyy}-${mm}-${dd}`; // Format: YYYY-MM-DD
  document.getElementById("dateAccessed").setAttribute("max", maxDate);
}

// Call the function when the Office add-in is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Set up the event handler for the submit button
    setMaxDate(); // Set the minimum date on load
    document.getElementById("submitButton").onclick = async function () {
      const acronym = document.getElementById("acronym").value;
      const yearPublished = document.getElementById("yearPublished").value;
      const title = document.getElementById("title").value;
      const url = document.getElementById("url").value;
      const dateAccessed = document.getElementById("dateAccessed").value;

      // Format the date
      const formattedDate = formatDate(dateAccessed);

      // Format the content to be inserted
      const formattedContent = `${acronym}. ${yearPublished}. ${title}. Retrieved from <${url}> (accessed on ${formattedDate}).\n`;

      await insertContent(formattedContent);
    };
  } else {
    // eslint-disable-next-line no-undef
    console.error("This add-in is not running in Word.");
  }
});

async function insertContent(content) {
  await Word.run(async (context) => {
    // Insert the formatted content at the current selection
    context.document.getSelection().insertText(content, Word.InsertLocation.end);
    await context.sync();
  });
}
function formatDate(dateString) {
  const date = new Date(dateString);
  const options = { month: "long", day: "numeric", year: "numeric" };
  return date.toLocaleDateString("en-US", options);
}
