/* global Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Office is ready!");

    // Hide sideload message
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    // Add button click handler
    document.getElementById("run").addEventListener("click", getSharePointData);
  }
});

async function getSharePointData() {
  OfficeRuntime.auth.getAccessToken({ forMSGraphAccess: true, allowSignInPrompt: true, allowConsentPrompt: true })
      .then(function(token) {
        console.log("Access token retrieved:", token);
        // Use the token for further authentication
      })
      .catch(function(error) {
        console.error("Error retrieving access token:", error);
      });
}
