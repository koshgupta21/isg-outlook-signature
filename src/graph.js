const GRAPH_API = "https://graph.microsoft.com/v1.0";

// Get access token from Office Single Sign-On
async function getAccessToken() {
    return new Promise((resolve, reject) => {
        Office.auth.getAccessToken({ allowSignInPrompt: true })
            .then(token => {
                console.log("✓ Access token obtained");
                resolve(token);
            })
            .catch(err => {
                console.error("✗ Failed to get access token:", err);
                reject(err);
            });
    });
}

// Fetch user profile from Microsoft Graph
async function getUserProfile() {
    try {
        console.log("Fetching user profile from Microsoft Graph...");

        const token = await getAccessToken();

        const response = await fetch(
            `${GRAPH_API}/me?$select=` +
            `displayName,jobTitle,mail,mobilePhone,businessPhones,officeLocation,department,onPremisesExtensionAttributes`,
            {
                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json"
                }
            }
        );

        if (!response.ok) {
            throw new Error(`Graph API error: ${response.status}`);
        }

        const data = await response.json();

        console.log("✓ User profile retrieved:", data.displayName);

        // AD Comments field syncs to extensionAttribute1
        // Confirm with your IT: might be extensionAttribute2, 3, etc.
        const comments = data.onPremisesExtensionAttributes?.extensionAttribute1 || "";

        return {
            displayName: data.displayName || "",
            jobTitle: data.jobTitle || "",
            email: (data.mail || "").toLowerCase(),
            mobilePhone: data.mobilePhone || "",
            businessPhone: (data.businessPhones || [])[0] || "",
            officeLocation: data.officeLocation || "",
            department: data.department || "",
            comments: comments
        };
    } catch (error) {
        console.error("Error fetching user profile:", error);
        throw error;
    }
}