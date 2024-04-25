// Listen for clicking Form Submit button
document.addEventListener("DOMContentLoaded", function () {
  const form = document.getElementById("capital-budget-form");
  form.addEventListener("submit", function (event) {
    event.preventDefault();
    const formData = new FormData(form);
    const data = {};
    for (const [key, value] of formData.entries()) {
      data[key] = value;
    }
    // Send Data to calculate new Google Sheet
    sendData(data);
  });

  function sendData(data) {
    fetch("/calculate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(data),
    })
      .then(() => {
        // Wait for 4 seconds
        setTimeout(function () {
          // reload the iframe to try to reload/refresh
          var iframe = document.getElementById("google-sheets-iframe");
          iframe.src = iframe.src;
        }, 4000);
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  }
});
// listen for clicking iframe to open in new window
document.getElementById("google-sheets").addEventListener("click", function () {
  window.open(
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vT5uLKvV_zEu_Bzs8rxI9bN0LbIb66Z7fO4ySPGSsX2pS8X6rTnDaST17c2__AA7uGQUMdmkY15ZxGM/pubhtml?gid=0&single=true"
  );
});
