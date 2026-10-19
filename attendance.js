export default async function handler(req, res) {
    // CORS Headers
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");

    if (req.method === "OPTIONS") {
        return res.status(200).end();
    }

    if (req.method !== "POST") {
        return res.status(455 || 405).json({
            success: false,
            error: "Method not allowed. Only POST requests are accepted."
        });
    }

    const appsScriptUrl = process.env.APPS_SCRIPT_URL;
    const apiSecret = process.env.APPS_SCRIPT_API_SECRET;

    if (!appsScriptUrl || !apiSecret) {
        return res.status(500).json({
            success: false,
            error: "Server configuration error: APPS_SCRIPT_URL or APPS_SCRIPT_API_SECRET missing in Vercel environment."
        });
    }

    try {
        const { action, token, attendeeIds, staffId } = req.body || {};

        // Action Whitelist Validation
        const allowedActions = ["inspectToken", "checkIn", "getDashboardSummary"];
        if (!action || !allowedActions.includes(action)) {
            return res.status(400).json({
                success: false,
                error: `Invalid action '${action}'. Allowed actions: ${allowedActions.join(", ")}`
            });
        }

        const payload = {
            action,
            apiSecret,
            token: token || "",
            attendeeIds: Array.isArray(attendeeIds) ? attendeeIds : [],
            staffId: staffId || ""
        };

        // Apps Script follows 302 redirects upon web app execution
        const response = await fetch(appsScriptUrl, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(payload),
            redirect: "follow"
        });

        const responseText = await response.text();

        let jsonResult;
        try {
            jsonResult = JSON.parse(responseText);
        } catch (parseErr) {
            return res.status(502).json({
                success: false,
                error: "Invalid JSON response received from Apps Script backend.",
                rawResponse: responseText.substring(0, 300)
            });
        }

        if (!jsonResult.success) {
            return res.status(400).json({
                success: false,
                error: jsonResult.error || "Backend execution failed."
            });
        }

        return res.status(200).json(jsonResult);
    } catch (err) {
        return res.status(500).json({
            success: false,
            error: "Internal server error connecting to Apps Script: " + (err.message || String(err))
        });
    }
}