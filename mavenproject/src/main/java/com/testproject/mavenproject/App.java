package com.testproject.mavenproject;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.concurrent.ThreadLocalRandom;

public class App {
    public static void main(String[] args) throws InterruptedException {
        Scanner scanner = new Scanner(System.in); // Keep scanner open for user input
        System.out.print("Enter the path to the Excel file: ");
        String excelFilePath = scanner.nextLine();
        
        System.setProperty("webdriver.chrome.driver", "resources\\chromedriver.exe"); 
        ChromeOptions options = new ChromeOptions();

        Map<String, Object> prefs = new HashMap<>();
        prefs.put("profile.default_content_setting_values.notifications", 2); 
        prefs.put("profile.default_content_setting_values.popups", 2); 
        prefs.put("profile.default_content_setting_values.cookies", 2); 
        prefs.put("profile.default_content_setting_values.automatic_downloads", 1); 
        options.setExperimentalOption("prefs", prefs);

        options.addArguments("--disable-popup-blocking"); 
        options.addArguments("--disable-blink-features=AutomationControlled"); 
        options.addArguments("--disable-translate");
        options.addArguments("--guest");

        WebDriver driver = new ChromeDriver(options);
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
             
            Sheet sheet = workbook.getSheetAt(0);
            
            Thread.sleep(ThreadLocalRandom.current().nextInt(2000, 2500));

            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) { 
                Row row = sheet.getRow(i);
                Cell domainCell = row.getCell(0); 
                Cell websiteTextCell = row.getCell(1); 

                if (domainCell == null || domainCell.getCellType() == CellType.BLANK) {
                    continue; 
                }

                String domain = "";
                if (domainCell.getCellType() == CellType.STRING) {
                    domain = domainCell.getStringCellValue();
                } else if (domainCell.getCellType() == CellType.NUMERIC) {
                    domain = String.valueOf(domainCell.getNumericCellValue());
                } else {
                    System.out.println("Non-string data in cell: " + domainCell.toString());
                }

                if (!domain.isEmpty()) {
                    try {
                        driver.get("https://www.google.com");
                        Thread.sleep(ThreadLocalRandom.current().nextInt(1500, 2000));
                        try {
                            WebElement searchBox = driver.findElement(By.name("q"));
                            searchBox.sendKeys("site:linkedin.com \"" + domain + "\"");
                            searchBox.submit();
                            
                            if (isCaptchaPresent(driver)) {
                                System.out.println("Captcha detected. Please complete the captcha and press Enter.");
                                scanner.nextLine(); 
                            }
                            
                            Thread.sleep(1500);
                            try {
                                List<WebElement> results = driver.findElements(By.xpath("//div[@class='MjjYud']//div[@data-snf='x5WNvb']//a"));
                                System.out.println(results.size() + " size of results ");
                                Thread.sleep(ThreadLocalRandom.current().nextInt(1000, 2000));                                
                                List<String> links = new ArrayList<>();
                                
                                for (WebElement result : results) {
                                    String link = result.getAttribute("href");
                                    if (link.contains("linkedin.com/company/") && !link.contains("translate.google.com")) {
                                        links.add(link);
                                        try {
                                            Document doc = Jsoup.connect(link).get();
                                            Element websiteDiv = doc.selectFirst("div[data-test-id=about-us__website]");
                                            String websiteText = "No website text found";
                                            if (websiteDiv != null) {
                                                Element websiteLink = websiteDiv.selectFirst("a[href]");
                                                if (websiteLink != null) {
                                                    websiteText = websiteLink.text();
                                                }
                                            }
                                            if (websiteTextCell == null) {
                                                websiteTextCell = row.createCell(1);
                                            }
                                            websiteTextCell.setCellValue(websiteText);
                                            if (!websiteText.equals("No website text found")) {
                                                System.out.println("Website Text: " + websiteText);
                                                break; 
                                            } else {
                                                System.out.println("No website text found for link: " + link);
                                            }
                                            System.out.println("Website Text: " + websiteText);
                                        } catch (IOException e) {
                                            e.printStackTrace();
                                        }
                                        System.out.println("Link added - " + link);
                                    }                            
                                }    
                            } catch (Exception e) {
                                System.out.println("There's no search results: " + e.getMessage());
                            }
                            
                        } catch (Exception e) {
                            System.out.println("Can't find search bar: " + e.getMessage());
                        }
                    } catch (Exception e) {
                        System.out.println("Can't access Google: " + e.getMessage());
                    }

                    // Write after each iteration to save progress
                    try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                        workbook.write(fos);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            driver.quit();
            scanner.close(); // Close scanner after usage
        }
    }

    private static boolean isCaptchaPresent(WebDriver driver) {
        try {
            // Check for common captcha elements or text
            WebElement captcha = driver.findElement(By.xpath("//*[contains(text(), 'captcha')]"));
            return captcha != null;
        } catch (Exception e) {
            return false;
        }
    }
}
