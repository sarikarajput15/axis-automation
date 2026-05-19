package com.insurance;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import technology.tabula.ObjectExtractor;
import technology.tabula.Page;
import technology.tabula.RectangularTextContainer;
import technology.tabula.Table;
import technology.tabula.extractors.SpreadsheetExtractionAlgorithm;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.regex.*;

/**
 * Java Playwright equivalent of task1.py
 * Automates insurance quote generation on Axis Max Life website.
 */
public class Task1 {

    // ── CONFIG ──────────────────────────────────────────────
    private static final String EXCEL_FILE = "../insurance_test_data.xlsx";
    private static final String OUTPUT_CSV = "../quote_results.csv";
    private static final int MAX_USERS = 5; // change this to run more

    // ── DEFAULTS (used when value missing/unknown) ───────────
    private static final Map<String, String> DEFAULTS = Map.ofEntries(
            Map.entry("gender", "Male"),
            Map.entry("tobacco", "No"),
            Map.entry("language", "English"),
            Map.entry("occupation", "Salaried"),
            Map.entry("education", "Graduate & Above"),
            Map.entry("diabetic", "No"),
            Map.entry("marital", "Single"),
            Map.entry("life_cover", "50 Lakhs"),
            Map.entry("cover_age", "60"),
            Map.entry("city", "Bangalore")
    );

    // ── MAPPINGS (Excel value → what the site shows) ─────────
    private static final Map<String, String> EDUCATION_MAP = Map.of(
            "grad or above", "Graduate & Above",
            "graduate & above", "Graduate & Above",
            "12th pass", "12",
            "10th pass", "10"
    );

    private static final Map<String, String> OCCUPATION_MAP = Map.of(
            "salaried", "Salaried",
            "self employed", "Self-employed/Business",
            "self-employed", "Self-employed/Business",
            "housewife", "Housewife"
    );

    // ── READ EXCEL ───────────────────────────────────────────
    private static List<Map<String, String>> loadUsers(String filepath, int maxUsers) throws IOException {
        List<Map<String, String>> users = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filepath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet ws = wb.getSheetAt(0);
            Iterator<Row> rowIterator = ws.iterator();

            // Read headers
            if (!rowIterator.hasNext()) return users;
            Row headerRow = rowIterator.next();
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(getCellValueAsString(cell).trim());
            }

            // Read data rows
            int count = 0;
            while (rowIterator.hasNext() && count < maxUsers) {
                Row row = rowIterator.next();
                Map<String, String> user = new LinkedHashMap<>();
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = row.getCell(i);
                    user.put(headers.get(i), cell != null ? getCellValueAsString(cell).trim() : "");
                }
                users.add(user);
                count++;
            }
        }
        return users;
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("dd/MM/yyyy");
                    yield sdf.format(cell.getDateCellValue());
                }
                double val = cell.getNumericCellValue();
                if (val == Math.floor(val) && !Double.isInfinite(val)) {
                    yield String.valueOf((long) val);
                }
                yield String.valueOf(val);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getStringCellValue();
            default -> "";
        };
    }

    // ── HELPERS ──────────────────────────────────────────────
    private static String get(Map<String, String> user, String key, String defaultValue) {
        String val = user.get(key);
        return (val != null && !val.isEmpty()) ? val.trim() : defaultValue;
    }

    private static String mapEducation(String val) {
        return EDUCATION_MAP.getOrDefault(val.toLowerCase(), DEFAULTS.get("education"));
    }

    private static String mapOccupation(String val) {
        return OCCUPATION_MAP.getOrDefault(val.toLowerCase(), DEFAULTS.get("occupation"));
    }

    /**
     * Convert '3 crore', '1.5 crore', '50 lakhs' → display label fragment.
     */
    private static String normalizeCover(String val) {
        val = val.toLowerCase().trim();
        if (val.contains("crore")) {
            String num = val.replace("crore", "").trim();
            return num.equals("1") ? "1 Crore" : num + " Crore";
        }
        if (val.contains("lakh") || val.contains("lac")) {
            Matcher m = Pattern.compile("[\\d.]+").matcher(val);
            if (m.find()) {
                return m.group() + " Lakhs";
            }
        }
        return val;
    }

    // ── EXTRACT 1ST YEAR PREMIUM FROM PDF ───────────────────
    private static String extractPdfPremium(String pdfPath) {
        try {
            File pdfFile = new File(pdfPath);
            try (PDDocument document = Loader.loadPDF(pdfFile)) {

                // ── Try structured table extraction first (using Tabula) ──
                try {
                    SpreadsheetExtractionAlgorithm sea = new SpreadsheetExtractionAlgorithm();
                    ObjectExtractor extractor = new ObjectExtractor(document);
                    for (int pageNum = 1; pageNum <= document.getNumberOfPages(); pageNum++) {
                        Page page = extractor.extract(pageNum);
                        List<Table> tables = sea.extract(page);
                        for (Table table : tables) {
                            for (int r = 0; r < table.getRowCount(); r++) {
                                List<RectangularTextContainer> row = table.getRows().get(r);
                                if (row.isEmpty()) continue;
                                String rowLabel = row.get(0).getText().trim().toLowerCase();
                                if (rowLabel.contains("first year gst")) {
                                    // Last non-empty cell = Total Installment Premium
                                    for (int c = row.size() - 1; c >= 0; c--) {
                                        String cellVal = row.get(c).getText().trim().replace(",", "");
                                        if (cellVal.matches("^\\d+(\\.\\d+)?$")) {
                                            return cellVal;
                                        }
                                    }
                                }
                            }
                        }
                    }
                } catch (Exception e) {
                    System.out.println("  ⚠ Tabula extraction failed, trying text: " + e.getMessage());
                }

                // ── Fallback: raw text scan ────────────────
                PDFTextStripper stripper = new PDFTextStripper();
                String text = stripper.getText(document);
                String[] lines = text.split("\n");
                for (String line : lines) {
                    if (line.toLowerCase().contains("first year gst")) {
                        Matcher m = Pattern.compile("[\\d,]+").matcher(line);
                        String lastMatch = null;
                        while (m.find()) {
                            lastMatch = m.group();
                        }
                        if (lastMatch != null) {
                            return lastMatch.replace(",", "");
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("  ⚠ PDF extraction error: " + e.getMessage());
        }
        return "N/A";
    }

    // ── SAVE RESULT ──────────────────────────────────────────
    private static void saveToCsv(Map<String, String> rowData) {
        String[] fieldnames = {
                "Test ID", "Insurer Name", "Equote Number",
                "Premium (1st Year)", "PDF Premium (1st Year)", "Premium Matched"
        };

        boolean fileExists = Files.exists(Path.of(OUTPUT_CSV));

        try (FileWriter fw = new FileWriter(OUTPUT_CSV, true);
             PrintWriter pw = new PrintWriter(fw)) {

            if (!fileExists) {
                pw.println(String.join(",", fieldnames));
            }

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < fieldnames.length; i++) {
                if (i > 0) sb.append(",");
                String value = rowData.getOrDefault(fieldnames[i], "");
                // Escape CSV values containing commas or quotes
                if (value.contains(",") || value.contains("\"") || value.contains("\n")) {
                    value = "\"" + value.replace("\"", "\"\"") + "\"";
                }
                sb.append(value);
            }
            pw.println(sb);

        } catch (IOException e) {
            System.out.println("  ⚠ CSV write error: " + e.getMessage());
        }
    }

    // ── EXTRACT RESULTS FROM PAGE ────────────────────────────
    private static Map<String, String> extractResults(com.microsoft.playwright.Page page, String insurerName) {
        page.waitForSelector("text=Review your Term Insurance summary",
                new com.microsoft.playwright.Page.WaitForSelectorOptions().setTimeout(15000));
        page.waitForTimeout(2000);

        String summaryText = page.innerText("body");
        String[] lines = summaryText.split("\n");
        List<String> cleanLines = new ArrayList<>();
        for (String line : lines) {
            String trimmed = line.trim();
            if (!trimmed.isEmpty()) {
                cleanLines.add(trimmed);
            }
        }

        // Equote — line immediately after "Equote Number"
        String equoteNumber = "N/A";
        for (int i = 0; i < cleanLines.size(); i++) {
            if (cleanLines.get(i).contains("Equote Number") && i + 1 < cleanLines.size()) {
                equoteNumber = cleanLines.get(i + 1).trim();
                break;
            }
        }

        // Premium — ₹ value after "Premium for 1st year"
        String premium = "N/A";
        for (int i = 0; i < cleanLines.size(); i++) {
            if (cleanLines.get(i).contains("Premium for 1st")) {
                for (int j = i; j < Math.min(i + 6, cleanLines.size()); j++) {
                    Matcher m = Pattern.compile("₹\\s*([\\d,]+\\.?\\d*)").matcher(cleanLines.get(j));
                    if (m.find()) {
                        premium = m.group(1).replace(",", "");
                        break;
                    }
                }
                break;
            }
        }

        String testId = UUID.randomUUID().toString().substring(0, 8).toUpperCase();

        Map<String, String> result = new LinkedHashMap<>();
        result.put("Test ID", testId);
        result.put("Insurer Name", insurerName);
        result.put("Equote Number", equoteNumber);
        result.put("Premium (1st Year)", premium);
        return result;
    }

    // ── MAIN FLOW PER USER ───────────────────────────────────
    private static void runForUser(Playwright playwright, Map<String, String> user) {
        String name = get(user, "Full Name", "Test User");
        String dob = get(user, "Date of birth", "01/01/1990");
        String mobile = get(user, "Mobile", "9999999999");
        String email = get(user, "email id", "test@gmail.com");
        String income = get(user, "Annual income", "1000000").replace(",", "");
        String pincode = get(user, "Pincode", "560001");
        String gender = get(user, "Gender", DEFAULTS.get("gender"));
        String occupation = mapOccupation(get(user, "Occupation", DEFAULTS.get("occupation")));
        String education = mapEducation(get(user, "Education", DEFAULTS.get("education")));
        String lifeCover = normalizeCover(get(user, "Life cover", DEFAULTS.get("life_cover")));
        String coverAge = get(user, "Cover till age", DEFAULTS.get("cover_age"));
        String rider = get(user, "Critical Illness Rider", "No");

        System.out.println("\n" + "=".repeat(50));
        System.out.println("Running for: " + name);
        System.out.println("  DOB: " + dob + " | Mobile: " + mobile + " | Gender: " + gender);
        System.out.println("  Occupation: " + occupation + " | Education: " + education);
        System.out.println("  Life Cover: " + lifeCover + " | Cover Age: " + coverAge + " yrs");
        System.out.println("  Rider: " + rider + " | Pincode: " + pincode);
        System.out.println("=".repeat(50));

        Browser browser = playwright.chromium().launch(new BrowserType.LaunchOptions().setHeadless(false));
        BrowserContext context = browser.newContext();
        com.microsoft.playwright.Page page = context.newPage();

        // Track the downloaded PDF path
        String downloadedPdfPath = null;

        try {
            // ── STEP 1: Personal details ────────────────────
            page.navigate("https://www.axismaxlife.com/term-insurance-plans/premium-calculator");
            page.getByRole(AriaRole.TEXTBOX, new com.microsoft.playwright.Page.GetByRoleOptions().setName("Full Name*")).fill(name);
            page.getByRole(AriaRole.TEXTBOX, new com.microsoft.playwright.Page.GetByRoleOptions().setName("Date of Birth*")).fill(dob);
            page.locator("#mobile").fill(mobile);
            page.locator("label").filter(new Locator.FilterOptions().setHasText("- 7")).click();
            page.getByRole(AriaRole.BUTTON, new com.microsoft.playwright.Page.GetByRoleOptions().setName("button")).click();
            page.waitForTimeout(3000);

            // ── STEP 2: Profile questions ───────────────────
            page.getByText(gender, new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).click();
            page.getByText("No", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).click();       // tobacco
            page.getByText("English", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).click();  // language
            page.getByText("Salaried", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).click();
            page.getByText("Graduate & Above").click();
            page.getByText("No", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).click();       // diabetic
            page.getByText("Single", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).click();   // marital
            page.getByText("Check Coverage", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).click();

            // ── STEP 3: Customize plan ──────────────────────
            page.waitForSelector("text=customize your Term Plan");
            page.waitForTimeout(2000);

            // Open life cover dropdown
            page.getByText(Pattern.compile("Recommended|Market Linked Returns")).first().click();
            page.waitForTimeout(1000);

            // Select life cover by partial match (with fallback to default if not found)
            try {
                Locator targetCover = page.locator("label, li")
                        .filter(new Locator.FilterOptions().setHasText(Pattern.compile(lifeCover, Pattern.CASE_INSENSITIVE)))
                        .first();
                targetCover.click(new Locator.ClickOptions().setTimeout(5000));
            } catch (Exception e) {
                String defaultLabel = normalizeCover(DEFAULTS.get("life_cover"));
                System.out.println("  ⚠ Life cover '" + lifeCover + "' not found, choosing default: " + defaultLabel);
                page.locator("label, li")
                        .filter(new Locator.FilterOptions().setHasText(Pattern.compile(defaultLabel, Pattern.CASE_INSENSITIVE)))
                        .first().click();
            }
            page.waitForTimeout(1000);

            page.getByText("MORE", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).first().click();
            try {
                page.getByRole(AriaRole.TEXTBOX, new com.microsoft.playwright.Page.GetByRoleOptions().setName("Enter your custom value")).fill(coverAge);
                page.getByText("Confirm", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).first().click();
            } catch (Exception e) {
                System.out.println("  ⚠ Could not fill custom value — skipping");
                try {
                    page.locator("a[role='button']:has-text('" + coverAge + " years')").click();
                    page.waitForTimeout(500);
                } catch (Exception e2) {
                    System.out.println("  ⚠ Could not find cover age option — skipping");
                }
            }

            page.getByText("Proceed", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).first().click();
            page.waitForTimeout(1000);

            // ── Critical Illness Rider ──────────────────────
            page.waitForTimeout(1500);
            try {
                if (rider.equalsIgnoreCase("yes")) {
                    page.evalOnSelector("#ci-rider input[name='riderAddButton']", "el => el.click()");
                    page.waitForTimeout(500);
                }
            } catch (Exception e) {
                System.out.println("  ⚠ Could not click Critical Illness rider — skipping");
            }

            try {
                page.locator("text=/Skip|Proceed/").first().click(new Locator.ClickOptions().setTimeout(5000));
                page.waitForTimeout(1000);
            } catch (Exception e) {
                // If neither is found, we assume the flow moved forward
            }

            // ── STEP 3b: Eligibility details ───────────────
            page.getByRole(AriaRole.TEXTBOX, new com.microsoft.playwright.Page.GetByRoleOptions().setName("Email Address*")).fill(email);
            page.getByRole(AriaRole.TEXTBOX, new com.microsoft.playwright.Page.GetByRoleOptions().setName("Annual Income*")).fill(income);
            page.getByRole(AriaRole.TEXTBOX, new com.microsoft.playwright.Page.GetByRoleOptions().setName("Pincode of Current Residential Address*")).fill(pincode);
            page.waitForTimeout(2000);

            // City dropdown
            page.locator("button.city-select").click();
            page.waitForTimeout(500);
            page.getByText("Bangalore", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).first().click();

            // ── Download Benefit Illustration & capture PDF ─
            Download download = page.waitForDownload(() -> {
                page.getByText("Download Benefit Illustration").click();
            });
            Path downloadPath = download.path();
            downloadedPdfPath = downloadPath != null ? downloadPath.toString() : null;
            System.out.println("  ✓ PDF downloaded to: " + downloadedPdfPath);

            page.waitForTimeout(8000);
            page.getByText("Proceed", new com.microsoft.playwright.Page.GetByTextOptions().setExact(true)).first().click();
            page.waitForTimeout(5000);

            // ── STEP 4: Extract from page ───────────────────
            Map<String, String> result = extractResults(page, name);

            // ── STEP 5: Extract PDF premium & compare ───────
            String pdfPremium = "N/A";
            String premiumMatched = "N/A";
            if (downloadedPdfPath != null) {
                pdfPremium = extractPdfPremium(downloadedPdfPath);
                if (!pdfPremium.equals("N/A") && !result.get("Premium (1st Year)").equals("N/A")) {
                    try {
                        int pageVal = (int) Double.parseDouble(result.get("Premium (1st Year)"));
                        int pdfVal = (int) Double.parseDouble(pdfPremium);
                        premiumMatched = Math.abs(pageVal - pdfVal) <= 1 ? "True" : "False";
                    } catch (NumberFormatException e) {
                        premiumMatched = "False";
                    }
                }
            }

            result.put("PDF Premium (1st Year)", pdfPremium);
            result.put("Premium Matched", premiumMatched);

            saveToCsv(result);

            System.out.println("  ✓ Test ID:       " + result.get("Test ID"));
            System.out.println("  ✓ Equote:        " + result.get("Equote Number"));
            System.out.println("  ✓ Page Premium:  ₹" + result.get("Premium (1st Year)"));
            System.out.println("  ✓ PDF Premium:   ₹" + pdfPremium);
            System.out.println("  ✓ Matched:       " + premiumMatched);
            System.out.println("  ✓ Saved to " + OUTPUT_CSV);

        } catch (Exception e) {
            System.out.println("  ✗ ERROR for " + name + ": " + e.getMessage());
            // Still save a row so we know it failed
            Map<String, String> errorRow = new LinkedHashMap<>();
            String errorMsg = e.getMessage();
            if (errorMsg != null && errorMsg.length() > 60) {
                errorMsg = errorMsg.substring(0, 60);
            }
            errorRow.put("Test ID", UUID.randomUUID().toString().substring(0, 8).toUpperCase());
            errorRow.put("Insurer Name", name);
            errorRow.put("Equote Number", "ERROR");
            errorRow.put("Premium (1st Year)", errorMsg != null ? errorMsg : "Unknown error");
            errorRow.put("PDF Premium (1st Year)", "N/A");
            errorRow.put("Premium Matched", "N/A");
            saveToCsv(errorRow);
        } finally {
            page.waitForTimeout(2000);
            context.close();
            browser.close();
        }
    }

    // ── ENTRY POINT ──────────────────────────────────────────
    public static void main(String[] args) {
        try {
            List<Map<String, String>> users = loadUsers(EXCEL_FILE, MAX_USERS);
            System.out.println("Loaded " + users.size() + " users from " + EXCEL_FILE);

            try (Playwright playwright = Playwright.create()) {
                for (int i = 0; i < users.size(); i++) {
                    System.out.println("\n[" + (i + 1) + "/" + users.size() + "] Starting...");
                    runForUser(playwright, users.get(i));
                }
            }

            System.out.println("\n✅ Done! Results saved to " + OUTPUT_CSV);

        } catch (Exception e) {
            System.err.println("Fatal error: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
