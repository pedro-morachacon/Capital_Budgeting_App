document.addEventListener("DOMContentLoaded", function () {
  const form = document.getElementById("capital-budget-form");
  form.addEventListener("submit", function (event) {
    event.preventDefault();
    const formData = new FormData(form);
    const data = {};
    for (const [key, value] of formData.entries()) {
      data[key] = value;
    }

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
        // reload the iframe
        var iframe = document.getElementById("google-sheets-iframe");
        iframe.src = iframe.src;
        // Redirect to the root page after submission
        window.location.href = "/";
        // reload the iframe
        var iframe = document.getElementById("google-sheets-iframe");
        iframe.src = iframe.src;
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  }
});
