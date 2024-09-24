document
  .getElementById("upload-form")
  .addEventListener("submit", function (event) {
    event.preventDefault();

    var fileInput = document.getElementById("pdf-file");
    var file = fileInput.files[0];

    if (file) {
      var formData = new FormData();
      formData.append("pdf", file);

      // Simulating processing time (remove this in actual implementation)
      setTimeout(function () {
        convertToExcel(formData);
      }, 3000);
    }
  });

function convertToExcel(formData) {
  fetch("http://172.30.29.33:8080/convert", {
    method: "POST",
    body: formData,
  })
    .then(function (response) {
      return response.blob();
    })
    .then(function (blob) {
      var downloadLink = document.createElement("a");
      downloadLink.href = URL.createObjectURL(blob);
      downloadLink.download = "output.xlsx";
      downloadLink.innerText = "Download Excel file";

      document.getElementById("result").innerHTML = "";
      document.getElementById("result").appendChild(downloadLink);
    })
    .catch(function (error) {
      console.error("Error:", error);
    });
}
