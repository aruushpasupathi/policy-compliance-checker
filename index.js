const { chromium } = require("playwright");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

// ---------------- CONFIG ----------------

const POLICY_CONFIG = {
  privacy: {
    anchorHints: ["privacy"],
    pageKeywords: ["privacy policy", "privacy notice", "privacy practices"],
    relevantFor: ["goods", "services"],
    displayName: "Privacy Policy"
  },
  terms: {
    anchorHints: ["terms"],
    pageKeywords: ["terms and conditions", "terms & conditions", "terms of use", "terms of service"],
    relevantFor: ["goods", "services"],
    displayName: "Terms and Conditions"
  },
  shipping: {
    anchorHints: ["shipping", "delivery"],
    pageKeywords: ["shipping", "delivery"],
    relevantFor: ["goods"],
    displayName: "Shipping Policy"
  },
  returns: {
    anchorHints: ["return", "returns"],
    pageKeywords: ["return policy", "returns policy", "return"],
    relevantFor: ["goods"],
    displayName: "Returns Policy"
  },
  refund: {
    anchorHints: ["refund"],
    pageKeywords: ["refund policy", "refund"],
    relevantFor: ["goods", "services"],
    displayName: "Refund Policy"
  },
  cancellation: {
    anchorHints: ["cancellation", "cancel"],
    pageKeywords: ["cancellation policy", "cancellation", "cancelling", "cancel"],
    relevantFor: ["goods", "services"],
    displayName: "Cancellation Policy"
  }
};

const POLICY_ORDER = {
  goods: ["privacy", "terms", "shipping", "returns", "refund", "cancellation"],
  services: ["privacy", "terms", "refund", "cancellation"]
};

const RESOLVABLE_GROUP = {
  goods: ["shipping", "returns", "refund", "cancellation"],
  services: ["refund", "cancellation"]
};

// ---------------- HELPERS ----------------

const normalize = s =>
  (s || "").toLowerCase().replace(/\s+/g, " ").trim();

const containsAny = (text, keywords) =>
  keywords.some(k => text.includes(k));

// Strict legal name matching
function containsLegalName(pageText, legalName) {
  if (!legalName || legalName.trim() === "") return false;
  
  const normalizedPageText = normalize(pageText);
  const normalizedLegalName = normalize(legalName);
  
  // Remove common suffixes
  const cleanLegalName = normalizedLegalName
    .replace(/\b(?:llc|inc|ltd|corp|corporation|co|company|limited)\b/g, "")
    .trim();
  
  if (!cleanLegalName) return false;
  
  // Strategy 1: Look for exact match with word boundaries
  const exactPattern = new RegExp(`\\b${cleanLegalName.replace(/\s+/g, '\\s+')}\\b`, 'i');
  if (exactPattern.test(pageText)) {
    return true;
  }
  
  // Strategy 2: Check if all name parts appear in close proximity
  const nameParts = cleanLegalName.split(/\s+/).filter(part => part.length > 2);
  
  if (nameParts.length >= 2) {
    // Check if at least 2 significant parts are found
    const foundParts = nameParts.filter(part => {
      const partPattern = new RegExp(`\\b${part}\\b`, 'i');
      return partPattern.test(pageText);
    });
    
    // Require at least 2 parts found for multi-part names
    if (foundParts.length >= 2) {
      // Verify parts appear in reasonable proximity
      const firstPartIndex = pageText.toLowerCase().indexOf(foundParts[0]);
      const lastPartIndex = pageText.toLowerCase().lastIndexOf(foundParts[foundParts.length - 1]);
      
      if (firstPartIndex !== -1 && lastPartIndex !== -1) {
        const distance = Math.abs(lastPartIndex - firstPartIndex);
        if (distance < 200) {
          return true;
        }
      }
    }
  } else if (nameParts.length === 1) {
    // For single-word names, use word boundary check
    const partPattern = new RegExp(`\\b${nameParts[0]}\\b`, 'i');
    return partPattern.test(pageText);
  }
  
  return false;
}

// Determine merchant type from various inputs
function determineMerchantType(input) {
  if (!input) return "goods"; // default
  const normalizedInput = normalize(input);
  // If "good" appears anywhere, treat as goods merchant
  if (normalizedInput.includes("good")) {
    return "goods";
  }
  // Otherwise treat as services
  return "services";
}

// Determine if entity is proprietorship
function isProprietorship(entityType) {
  if (!entityType) return true; // default to proprietorship for safety
  const normalizedEntity = normalize(entityType);
  return normalizedEntity.includes("proprietor");
}

async function safeGoto(page, url) {
  try {
    await page.goto(url, {
      waitUntil: "domcontentloaded",
      timeout: 45000
    });
    return true;
  } catch {
    return false;
  }
}

// Check if a policy is relevant for the merchant type
function isPolicyRelevant(policy, merchantType) {
  return POLICY_CONFIG[policy].relevantFor.includes(merchantType);
}

// Generate email content based on missing policies
function generateEmailContent(website, legalName, merchantEmail, missingPolicies, entityType) {
  const isProprietor = isProprietorship(entityType);
  
  // Separate legal name from other missing policies
  const hasLegalNameMissing = missingPolicies.includes("legal name");
  const otherMissingPolicies = missingPolicies.filter(p => p !== "legal name");
  
  let emailBody = `Dear ${legalName || "Merchant"},\n\n`;
  emailBody += `Following our initial review for your PayGlocal onboarding, we need your assistance to ensure your website meets our required compliance standards.\n\n`;
  emailBody += `Please implement the following updates on ${website}:\n\n`;
  
  // Add legal name section if missing
  if (hasLegalNameMissing && isProprietor) {
    emailBody += `1. As the company has been registered as a proprietorship, RBI guidelines require you to display your legal name on the website. Please update the website to display your registered legal name: ${legalName}\n\n`;
    emailBody += `   Possible locations to display your legal name:\n`;
    emailBody += `   • Website Footer (Most common)\n`;
    emailBody += `   • About Us page\n`;
    emailBody += `   • Contact Us page\n`;
    emailBody += `   • Site-Wide copyright notice (e.g., "© ${new Date().getFullYear()} ${legalName}. All rights reserved.")\n\n`;
  }
  
  // Add policies section if missing
  if (otherMissingPolicies.length > 0) {
    if (hasLegalNameMissing) {
      emailBody += `2. `;
    }
    emailBody += `It is required that you add the following policies to your website for legal protection:\n`;
    emailBody += `   • ${otherMissingPolicies.map(policy => {
      // Convert policy key to display name
      if (policy === "privacy") return "Privacy Policy";
      if (policy === "terms") return "Terms and Conditions";
      if (policy === "shipping") return "Shipping Policy";
      if (policy === "returns") return "Returns Policy";
      if (policy === "refund") return "Refund Policy";
      if (policy === "cancellation") return "Cancellation Policy";
      return policy;
    }).join("\n   • ")}\n\n`;
  }
  
  emailBody += `These updates are necessary before we can activate your account. Once completed, kindly notify us so we can verify and proceed.\n\n`;
  emailBody += `Thank you for your prompt attention to this.\n\n`;
  emailBody += `Sincerely,\nPayGlocal Onboarding Team`;
  
  return {
    subject: "PayGlocal Onboarding: Website Updates Needed",
    to: merchantEmail,
    body: emailBody
  };
}

// Create email directory if it doesn't exist
function ensureEmailDirectory() {
  const emailDir = path.join(process.cwd(), "emails");
  if (!fs.existsSync(emailDir)) {
    fs.mkdirSync(emailDir, { recursive: true });
  }
  return emailDir;
}

// Save email to file 
function saveEmailToFile(website, merchantEmail, emailContent, index) {
  const emailDir = ensureEmailDirectory();
  
  // Create a safe filename from website
  const safeWebsiteName = website
    .replace(/^https?:\/\//, '')
    .replace(/[^a-z0-9]/gi, '_')
    .toLowerCase();
  
  const filename = `email_${index + 1}_${safeWebsiteName}.txt`;
  const filepath = path.join(emailDir, filename);
  
  // Build the email content string
  const fullEmail = `To: ${merchantEmail}\nSubject: ${emailContent.subject}\n\n${emailContent.body}`;
  
  fs.writeFileSync(filepath, fullEmail, 'utf8');
  return filename;
}

// ---------------- CHECK SINGLE WEBSITE ----------------

async function checkWebsite(page, website, merchantTypeInput, entityType, legalNameRaw) {
  // Determine merchant type from input
  const merchantType = determineMerchantType(merchantTypeInput);
  const legalName = legalNameRaw || "";
  const ORDER = POLICY_ORDER[merchantType];
  const GROUP = RESOLVABLE_GROUP[merchantType];
  
  // Check if entity is proprietorship
  const isProprietorEntity = isProprietorship(entityType);

  const result = {
    website,
    merchantType: merchantTypeInput || "", // Keep original input
    determinedMerchantType: merchantType, // Store determined type
    entityType: entityType || "",
    isProprietorship: isProprietorEntity,
    legalName: legalName,
    legalNamePresent: isProprietorEntity ? false : "NOT RELEVANT",
    legalNameURL: null,
    policies: {},
    missingPolicies: [],
    allRelevantPoliciesMissing: false,
    complianceStatus: "FAIL",
    error: null
  };

  // Initialize all policies
  Object.keys(POLICY_CONFIG).forEach(p => {
    if (isPolicyRelevant(p, merchantType)) {
      result.policies[p] = { 
        present: false, 
        url: null, 
        status: "MISSING" // Will be updated
      };
    } else {
      result.policies[p] = { 
        present: false, 
        url: null, 
        status: "NOT RELEVANT" 
      };
    }
  });

  try {
    const ok = await safeGoto(page, website);
    if (!ok) {
      result.error = "Failed to load website";
      return result;
    }

    // ---------------- EXTRACT LINKS ----------------
    const rawLinks = await page.$$eval("a", as =>
      as.map(a => ({
        text: a.innerText || "",
        href: a.href || ""
      }))
    );

    const links = rawLinks.map(l => ({
      text: normalize(l.text),
      href: l.href
    }));

    // Store visited pages for comprehensive checking
    const visitedPages = [];

    // ---------------- PRIMARY PASS (ORDERED) ----------------
    for (const policy of ORDER) {
      // Skip if policy not relevant
      if (!isPolicyRelevant(policy, merchantType)) continue;
      
      for (const link of links) {
        if (!link.href || !link.text) continue;
        if (!containsAny(link.text, POLICY_CONFIG[policy].anchorHints)) continue;

        const ok = await safeGoto(page, link.href);
        if (!ok) continue;

        const pageText = normalize(await page.textContent("body"));
        
        // Store this page
        visitedPages.push({
          url: link.href,
          pageText: pageText,
          policyType: policy
        });

        if (containsAny(pageText, POLICY_CONFIG[policy].pageKeywords)) {
          result.policies[policy] = {
            present: true,
            url: link.href,
            pageText,
            status: "FOUND"
          };

          // Check for legal name in policy text for proprietorships
          if (
            isProprietorEntity &&
            result.legalNamePresent === false &&
            legalName &&
            containsLegalName(pageText, legalName)
          ) {
            result.legalNamePresent = true;
            result.legalNameURL = link.href;
          }

          // Navigate back to homepage for next link check
          await safeGoto(page, website);
          break;
        }
      }
    }

    // ---------------- GROUP RESOLUTION ----------------
    for (let i = 0; i < GROUP.length; i++) {
      const presentSources = GROUP
        .map(p => result.policies[p])
        .filter(p => p && p.present && p.pageText);

      if (presentSources.length === 0) break;

      for (const policy of GROUP) {
        // Skip if policy not relevant or already found
        if (!isPolicyRelevant(policy, merchantType) || result.policies[policy].present) continue;

        for (const source of presentSources) {
          if (containsAny(source.pageText, POLICY_CONFIG[policy].pageKeywords)) {
            result.policies[policy] = {
              present: true,
              url: source.url,
              pageText: source.pageText,
              status: "FOUND"
            };
            
            // Also check for legal name in this policy text
            if (
              isProprietorEntity &&
              result.legalNamePresent === false &&
              legalName &&
              containsLegalName(source.pageText, legalName)
            ) {
              result.legalNamePresent = true;
              result.legalNameURL = source.url;
            }
            
            break;
          }
        }
      }
    }

    // ---------------- VERIFICATION STEP ----------------
    // If legal name was supposedly found, verify it's actually there
    if (isProprietorEntity && result.legalNamePresent === true && result.legalNameURL) {
      const verifyOk = await safeGoto(page, result.legalNameURL);
      if (verifyOk) {
        const verifyText = normalize(await page.textContent("body"));
        if (!containsLegalName(verifyText, legalName)) {
          // False positive! Reset and continue searching
          result.legalNamePresent = false;
          result.legalNameURL = null;
        }
      }
    }

    // ---------------- COMPREHENSIVE SEARCH ----------------
    if (isProprietorEntity && result.legalNamePresent === false && legalName) {
      // 1. Check all visited pages first
      for (const pageInfo of visitedPages) {
        if (containsLegalName(pageInfo.pageText, legalName)) {
          result.legalNamePresent = true;
          result.legalNameURL = pageInfo.url;
          break;
        }
      }
      
      // 2. Check homepage
      if (result.legalNamePresent === false) {
        await safeGoto(page, website);
        const homepageText = normalize(await page.textContent("body"));
        if (containsLegalName(homepageText, legalName)) {
          result.legalNamePresent = true;
          result.legalNameURL = website;
        }
      }
      
      // 3. Check other common pages
      if (result.legalNamePresent === false) {
        const otherPages = ["contact", "about", "info", "us", "company"];
        for (const pageType of otherPages) {
          const relevantLinks = links.filter(link => 
            link.text.includes(pageType) && link.href
          );
          
          for (const link of relevantLinks.slice(0, 2)) {
            const ok = await safeGoto(page, link.href);
            if (!ok) continue;
            
            const pageText = normalize(await page.textContent("body"));
            if (containsLegalName(pageText, legalName)) {
              result.legalNamePresent = true;
              result.legalNameURL = link.href;
              break;
            }
            
            await safeGoto(page, website);
            if (result.legalNamePresent === true) break;
          }
          if (result.legalNamePresent === true) break;
        }
      }
    }

    // ---------------- FINAL EVALUATION ----------------
    for (const policy of ORDER) {
      if (!isPolicyRelevant(policy, merchantType)) continue;
      
      if (!result.policies[policy].present) {
        result.missingPolicies.push(policy);
        result.policies[policy].status = "MISSING";
      }
      // Clean up pageText from result object
      if (result.policies[policy].pageText) {
        delete result.policies[policy].pageText;
      }
    }

    // Add "legal name" to missing policies if it's a proprietorship and legal name not found
    if (isProprietorEntity && result.legalNamePresent === false && legalName) {
      result.missingPolicies.push("legal name");
    }

    // Check if all relevant policies are missing
    const allRelevantPolicies = ORDER.filter(policy => isPolicyRelevant(policy, merchantType));
    const foundRelevantPolicies = allRelevantPolicies.filter(policy => 
      result.policies[policy].present === true
    );
    
    // If no relevant policies were found at all
    result.allRelevantPoliciesMissing = foundRelevantPolicies.length === 0;

    // Check compliance
    const allRelevantPoliciesPresent = ORDER.every(policy => 
      !isPolicyRelevant(policy, merchantType) || result.policies[policy].present
    );
    
    const legalNameOk = !isProprietorEntity || result.legalNamePresent === true;
    
    if (allRelevantPoliciesPresent && legalNameOk) {
      result.complianceStatus = "PASS";
    }

  } catch (error) {
    result.error = error.message;
  }

  return result;
}

// ---------------- MAIN ----------------

(async () => {
  try {
    // Read input file
    const inputFile = "input.xlsx"; // This must be the input file name
    let workbook;
    try {
      workbook = XLSX.readFile(inputFile);
    } catch (error) {
      console.error(`Error: Cannot read input.xlsx file. Please make sure it exists in the same folder as index.js`);
      console.error(`Current directory: ${process.cwd()}`);
      console.error(`Looking for: ${inputFile}`);
      process.exit(1);
    }
    
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to JSON
    const data = XLSX.utils.sheet_to_json(worksheet);
    
    console.log(`Processing ${data.length} websites...`);
    
    // Launch browser once
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage();
    
    // Create array to store email information
    const emailsToSend = [];
    
    // Process each row
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const website = row.Website || row.website;
      
      if (!website) {
        console.log(`Skipping row ${i + 1}: No website URL`);
        continue;
      }
      
      console.log(`Processing ${i + 1}/${data.length}: ${website}`);
      
      // Extract data from Excel columns with exact column names
      const merchantType = row.MerchantType || row.merchantType || "goods";
      const entityType = row.EntityType || row.entityType || "proprietorship";
      const legalName = row.LegalName || row.legalName || "";
      const merchantEmail = row.Email || row.email || "";
      
      // Check the website
      const result = await checkWebsite(
        page, 
        website, 
        merchantType, 
        entityType, 
        legalName
      );
      
      // Add results to the row
      data[i].ComplianceStatus = result.complianceStatus;
      data[i].MissingPolicies = result.missingPolicies.join(", ");
      data[i].LegalNamePresent = result.legalNamePresent;
      data[i].IsProprietorship = result.isProprietorship;
      data[i].DeterminedMerchantType = result.determinedMerchantType;
      data[i].Error = result.error;
      
      // Add ManualCheckingRequired column
      data[i].ManualCheckingRequired = result.allRelevantPoliciesMissing ? "YES" : "NO";
      
      // Add LegalNameStatus column 
      if (result.isProprietorship) {
        data[i].LegalNameStatus = result.legalNamePresent === true ? "FOUND" : "MISSING";
      } else {
        data[i].LegalNameStatus = "NOT RELEVANT";
      }
      
      // Add LegalNameURL column if legal name was found
      if (result.legalNameURL) {
        data[i].LegalNameURL = result.legalNameURL;
      }
      
      // Add individual policy status with URLs
      Object.keys(POLICY_CONFIG).forEach(policy => {
        const policyResult = result.policies[policy];
        data[i][`${policy.charAt(0).toUpperCase() + policy.slice(1)}Status`] = policyResult.status;
        if (policyResult.url) {
          data[i][`${policy.charAt(0).toUpperCase() + policy.slice(1)}URL`] = policyResult.url;
        }
      });
      
      // Generate email if conditions are met
      if (result.complianceStatus === "FAIL" && 
          !result.allRelevantPoliciesMissing && 
          merchantEmail && 
          result.missingPolicies.length > 0) {
        
        const emailContent = generateEmailContent(
          website,
          legalName,
          merchantEmail,
          result.missingPolicies,
          entityType
        );
        
        // Save email to file
        const emailFilename = saveEmailToFile(website, merchantEmail, emailContent, i);
        
        // Store email info for summary
        emailsToSend.push({
          website,
          merchantEmail,
          filename: emailFilename,
          missingPolicies: result.missingPolicies
        });
        
        data[i].EmailGenerated = "YES";
        data[i].EmailFilename = emailFilename;
      } else {
        data[i].EmailGenerated = "NO";
        data[i].EmailFilename = "";
      }
      
      // Delay between requests
      await new Promise(resolve => setTimeout(resolve, 2000));
    }
    
    await browser.close();
    
    // Convert back to worksheet
    const newWorksheet = XLSX.utils.json_to_sheet(data);
    
    // Create new workbook with results
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Results");
    
    // Save to output file
    const outputFile = "output.xlsx";
    XLSX.writeFile(newWorkbook, outputFile);
    
    console.log(`\nDone! Results saved to ${outputFile}`);
    
    // Print summary
    const manualCheckSites = data.filter(row => row.ManualCheckingRequired === "YES");
    if (manualCheckSites.length > 0) {
      console.log(`\n=== MANUAL CHECKING REQUIRED (${manualCheckSites.length} sites) ===`);
      manualCheckSites.forEach((row, index) => {
        console.log(`${index + 1}. ${row.Website} - No relevant policies found`);
      });
      console.log(`Please manually check these websites as the crawler couldn't find any policies.`);
    }
    
    // Print email summary
    if (emailsToSend.length > 0) {
      console.log(`\n=== EMAILS GENERATED (${emailsToSend.length} sites) ===`);
      console.log(`Emails have been saved to the "emails" folder.`);
      console.log(`\nSummary of emails to send:`);
      emailsToSend.forEach((email, index) => {
        console.log(`${index + 1}. ${email.website}`);
        console.log(`   To: ${email.merchantEmail}`);
        console.log(`   Missing: ${email.missingPolicies.join(", ")}`);
        console.log(`   File: emails/${email.filename}`);
        console.log(``);
      });
      
      // Email Summary
      const emailSummaryData = emailsToSend.map(email => ({
        Website: email.website,
        Email: email.merchantEmail,
        MissingPolicies: email.missingPolicies.join(", "),
        EmailFile: `emails/${email.filename}`,
        Status: "Ready to Send"
      }));
      
      const emailSummarySheet = XLSX.utils.json_to_sheet(emailSummaryData);
      XLSX.utils.book_append_sheet(newWorkbook, emailSummarySheet, "Email Summary");
      XLSX.writeFile(newWorkbook, outputFile); 
      
      console.log(`Email summary has been added to the "Email Summary" sheet in ${outputFile}`);
    } else {
      console.log(`\nNo emails were generated (all sites either passed, require manual checking, or have no email).`);
    }
    
  } catch (error) {
    console.error("Error:", error);
    process.exit(1);
  }
})();