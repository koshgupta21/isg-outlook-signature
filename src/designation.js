// Port of GetDesignation() from C# to JavaScript
function getDesignation(comments) {
    if (!comments) return "";

    const c = comments.toLowerCase();

    // Check for codes in order
    if (c.includes("daiadv"))   return "AI Certified | Advanced";
    if (c.includes("exa") ||
        c.includes("exadap") ||
        c.includes("exapp")  ||
        c.includes("excybe") ||
        c.includes("exclou") ||
        c.includes("exengi") ||
        c.includes("exwrk")  ||
        c.includes("exdata"))   return "AI Certified | Expert";
    if (c.includes("daif"))     return "AI Certified | Fellow";
    if (c.includes("26adv"))    return "AI Cert 2026 Advanced";
    if (c.includes("26aorg"))   return "AI Cert 2026 - Adaptive Orgs Expert";
    if (c.includes("26cyb"))    return "AI Cert 2026 - Cybersecurity Expert";
    if (c.includes("26data"))   return "AI Cert 2026 - Data & Applied AI Expert";
    if (c.includes("26eng"))    return "AI Cert 2026 - Engineering Expert";
    if (c.includes("26fel"))    return "AI Cert 2026 Fellow";
    if (c.includes("26tech"))   return "AI Cert 2026 - Tech Mod";

    return "";
}

// Returns CDN URL for badge image, or empty string if no badge
function getDesignationBadgeUrl(designation) {
    // Replace with your actual CDN URL
    const CDN_BASE = "https://cdn.isg-one.com/Email%20Images";

    if (!designation) return "";

    if (designation === "AI Cert 2026 Advanced")
        return `${CDN_BASE}/AI_Certified_Badge_2026_Advanced_Dark_3X.png`;

    if (designation === "AI Cert 2026 Fellow")
        return `${CDN_BASE}/AI_Certified_Badge_2026_Fellow_Dark_3X.png`;

    if (designation.startsWith("AI Cert 2026 -"))
        return `${CDN_BASE}/AI_Certified_Badge_2026_Expert_Dark_3X.png`;

    return "";
}

// Check if user is Chairman's Club
function isChairmansClub(comments) {
    if (!comments) return false;
    const c = comments.toLowerCase();
    return c.includes("cc") || c.includes("chairman");
}