package com.example.testcallchrome;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static com.codeborne.selenide.Selenide.sleep;

public class MainPage {

    public static void checkLogin(WebDriver webDriver) {
        boolean login = false;
        while (!login) {
            WebElement loginBtn = null;
            try {
                loginBtn = webDriver.findElement(By.cssSelector("div.nologin.float_r.cancle_select_bg"));
                System.out.println("未登录...");
            } catch (Exception e) {

            }

            try {
                WebElement userInfo = webDriver.findElement(By.cssSelector("div.header_wrap > div.lmzHeader > div > div.person_msg.float_r > span:nth-child(1)"));
                if (userInfo != null && !userInfo.getText().isEmpty()) {
                    System.out.println("已登录...");
                    login = true;
                } else {
                    System.out.println("登录状态未知...");
                    login = true;
                }
            } catch (Exception e) {

            }

            if (loginBtn != null) {
                try {
                    loginBtn.click();
                } catch (Exception e) {
                }
                sleep(5000);
            }

            sleep(1000);
        }
    }

    public static void main(String[] args) {
        if (isWindows()) {
            System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
        } else {
            System.setProperty("webdriver.chrome.driver", "chromedriver");
        }

        ChromeOptions chromeOptions = new ChromeOptions();

        String username = System.getProperty("user.name");
        System.out.println("=======================");
        System.out.println("hello " + username);
        System.out.println("=======================");

        if (isWindows()) {
            chromeOptions.addArguments("user-data-dir=C:/Users/" + username + "/AppData/Local/Google/Chrome/User Data/Default");
        } else {
            chromeOptions.addArguments("user-data-dir=/Users/" + username + "/Library/Application Support/Google/Chrome/");

        }

        WebDriver webDriver = new ChromeDriver(chromeOptions);

        while (true) {
            webDriver.get("https://www.gaokao.cn/school/search");
            System.out.println("加载首页...");
            sleep(1000);

            checkLogin(webDriver);

            // start process
            // get school list

            int processCnt = 0;

            List<WebElement> schoolTrList = webDriver.findElements(By.cssSelector("#myTable > tbody > tr"));
            for (WebElement schoolTr : schoolTrList) {
                WebElement schoolIcon = schoolTr.findElement(By.cssSelector("td.school-icon > a > img"));
                WebElement schoolName = schoolTr.findElement(By.cssSelector("td.name-des > div.top-item > a > span.float_l.set_hoverl.am_l"));
                String schoolHref = schoolIcon.getAttribute("src");
                String schoolNameStr = schoolName.getText();
                String schoolIndex = schoolHref.substring(schoolHref.lastIndexOf("/") + 1, schoolHref.lastIndexOf("."));
                System.out.println(schoolIndex + "_" + schoolNameStr);

                if (Files.exists(Path.of(schoolNameStr + ".xls"))) {
                    System.out.println("file exist " + Path.of(schoolNameStr + ".xls"));
                    continue;
                }

                startProcess(webDriver, schoolNameStr, schoolIndex);
                processCnt++;
                break;
            }

            if (processCnt == 0) {
                try {
                    WebElement nextPageBtn = webDriver.findElement(By.cssSelector("#root i.anticon.anticon-right"));
                    nextPageBtn.click();
                } catch (Exception e) {
                    System.out.println("no data end...");
                    break;
                }
            }
        }

        webDriver.close();
    }

    static boolean isWindows() {
        return System.getProperties().getProperty("os.name").toUpperCase().contains("WINDOWS");
    }

    static void startProcess(WebDriver webDriver, String schoolName, String schoolIndex) {
        webDriver.get("https://gkcx.eol.cn/school/" + schoolIndex + "/provinceline");
        sleep(300);
        String title = webDriver.getTitle();
        //let childCount = document.querySelector("#b5776711-83dc-4f80-84ad-171e94a4f5bf > ul").childElementCount;

        //

        List<WebElement> dropDownList = webDriver.findElements(By.cssSelector("#scoreline > div.content-top.content_top_1_4 > div.scoreLine-dropDown > div.dropdown-box"));

        WebElement provinceElement = webDriver.findElement(By.cssSelector("#scoreline > div.content-top.content_top_1_4 > div > div:nth-child(1) > div > div > div > div"));
        sleep(50);
        provinceElement.click();
        sleep(150);

        checkLogin(webDriver);

//        WebElement element2 = webDriver.findElement(By.cssSelector("#scoreline > div.content-top.content_top_1_4 > div > div:nth-child(2)"));
//        element2.click();
//        System.out.println(element2.getText());

        List<WebElement> provinceList = webDriver.findElements(By.cssSelector("ul[role=listbox] > li"));

        Data data = new Data();
        data.school = schoolName;
        data.tableTitle = dropDownList.stream().map(WebElement::getText).collect(Collectors.joining("-"));

        Map<String, List<List<String>>> datas = new LinkedHashMap<>();
        data.datas = datas;

        for (int i = 0; i < provinceList.size(); ) {
            try {
                WebElement province = provinceList.get(i);
                sleep(50);
                provinceElement.click();
                sleep(180);
                province.click();
                sleep(220);

                System.out.println("=============" + provinceElement.getText() + "=============");
                List<List<String>> d = new ArrayList<>();
                datas.put(provinceElement.getText(), d);

                // th
                List<WebElement> ths = webDriver.findElements(By.cssSelector("#scoreline > div.province_score_line_table > div.line_table_box.major_score_table > table > tbody > tr:nth-child(1) > th"));
                List<String> dh = new ArrayList<>();
                d.add(dh);
                for (WebElement th : ths) {
                    System.out.print(th.getText());
                    System.out.print("\t");
                    dh.add(th.getText());
                }

                while (true) {
                    WebElement nextPage = null;
                    try {
                        nextPage = webDriver.findElement(By.cssSelector("#scoreline > div.province_score_line_table > div.table_pagination_box > div > div > ul > li.ant-pagination-next"));
                    } catch (Exception e) {
                        // no next page
                    }

                    // td
                    List<WebElement> trs = webDriver.findElements(By.cssSelector("#scoreline > div.province_score_line_table > div.line_table_box.major_score_table > table > tbody > tr"));
                    for (int j = 0; j < trs.size(); j++) {
                        List<String> dd = new ArrayList<>();

                        WebElement tr = trs.get(j);
                        List<WebElement> tds = new ArrayList<>();
                        try {
                            tds = tr.findElements(By.cssSelector("td"));
                        } catch (Exception e) {
                        }
                        for (int k = 0; k < tds.size(); k++) {
                            System.out.print(tds.get(k).getText());
                            System.out.print("\t");
                            dd.add(tds.get(k).getText());
                        }
                        System.out.println("");

                        if (dd.stream().allMatch(String::isEmpty)) {
                            // remove empty
                            continue;
                        }

                        d.add(dd);
                    }

                    if (nextPage == null || nextPage.getAttribute("class").contains("ant-pagination-disabled")) {
                        break;
                    } else {
                        sleep(50);
                        nextPage.click();
                        sleep(150);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
                continue;
            }

            i++;
        }

        createReport(data);
    }

    static class Data {
        String tableTitle;
        String school;
        Map<String, List<List<String>>> datas;
    }

    private static File createReport(Data data) {
        File file = null;
        try {
            file = Files.createFile(Path.of(data.school + ".xls")).toFile();  //临时文件
            System.out.println(file.getAbsolutePath());
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(data.school);

        data.datas.forEach((province, datas) -> {
            HSSFRow nameRow = sheet.createRow(sheet.getLastRowNum() + 1);
            nameRow.createCell(0).setCellValue(province + " " + data.tableTitle);

            for (int i = 0; i < datas.size(); i++) {
                List<String> d = datas.get(i);
                if (d.isEmpty()) {
                    continue;
                }
                HSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
                if (i == 0) {
                    HSSFCellStyle style = workbook.createCellStyle();
                    style.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
                    row.setRowStyle(style);
                }

                for (int j = 0; j < d.size(); j++) {
                    row.createCell(j).setCellValue(d.get(j));
                }
            }

            HSSFRow endRow = sheet.createRow(sheet.getLastRowNum() + 1);
            endRow.createCell(0).setCellValue("");
            endRow = sheet.createRow(sheet.getLastRowNum() + 1);
            endRow.createCell(0).setCellValue("");
        });

        try {
            FileOutputStream fileOut = new FileOutputStream(file);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            System.out.println("保存excel success！");
        } catch (Exception e) {
            e.printStackTrace();
        }
        return file;
    }
}
