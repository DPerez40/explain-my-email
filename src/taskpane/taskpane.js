Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run").onclick = async () => {
      const item = Office.context.mailbox.item;
      const graphItemId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);

      try {
        // STEP 1: Get Graph access token
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const accessToken = result.value;

            // STEP 2: Fetch email data from Microsoft Graph
            const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}`, {
              method: "GET",
              headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Accept": "application/json"
              }
            });

            const data = await response.json();

            const sender = data.from?.emailAddress?.address || "Unknown Sender";
            const subject = data.subject || "No Subject";
            const bodyPreview = data.bodyPreview || "No Body Preview";

            // Get InternetMessageHeaders (contains all headers like SPF, DKIM, etc.)
            const headersResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphItemId}/internetMessageHeaders`, {
              method: "GET",
              headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Accept": "application/json"
              }
            });

            const headersData = await headersResponse.json();
            const headersList = headersData.value.map(h => `${h.name}: ${h.value}`).join("\n");

            // Display it all
            document.getElementById("status").innerText =
              `From: ${sender}\nSubject: ${subject}\n\n--- HEADERS ---\n${headersList}\n\n--- Body Preview ---\n${bodyPreview}`;
          } else {
            document.getElementById("status").innerText = "Failed to get access token.";
          }
        });
      } catch (error) {
        console.error(error);
        document.getElementById("status").innerText = "An error occurred.";
      }
    };
  }
});
