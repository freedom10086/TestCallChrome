package com.example.testcallchrome;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
import java.util.List;
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
            chromeOptions.addArguments("user-data-dir=/Users/" + username + "/Library/Application Support/Google/Chrome/Default");

        }

        WebDriver webDriver = new ChromeDriver(chromeOptions);

        while (true) {
            if (!"https://www.gaokao.cn/school/search".equals(webDriver.getCurrentUrl())) {
                webDriver.get("https://www.gaokao.cn/school/search");
                System.out.println("加载首页...");
                sleep(1000);
            }
            System.out.println("current page url:" + webDriver.getCurrentUrl());

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
                    System.out.println("page all done try find next page");
                    WebElement nextPageBtn = webDriver.findElement(By.cssSelector("#root i.anticon.anticon-right > svg"));
                    nextPageBtn.click();
                    System.out.println("click nextPageBtn!");
                } catch (Exception e) {
                    e.printStackTrace();
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

        List<WebElement> dropDownList = webDriver.findElements(By.cssSelector("#scoreline > div.content-top.content_top_1_4 > div.scoreLine-dropDown > div.dropdown-box"));

        sleep(50);
        dropDownList.forEach(e -> {
            try {
                e.click();
                System.out.println("click " + e.getText());
            } catch (Exception ee) {

            }
            sleep(150);
        });

        checkLogin(webDriver);

        Data data = new Data();
        data.school = schoolName;
        data.datas = getAllTableData(webDriver, 0);

        createReport(data);
    }

    static List<List<String>> getAllTableData(WebDriver webDriver, int index) {
        List<List<String>> datas = new ArrayList<>();

        List<WebElement> dropDownList = webDriver.findElements(By.cssSelector("#scoreline > div.content-top.content_top_1_4 > div.scoreLine-dropDown > div.dropdown-box"));
        WebElement dropDownItem = dropDownList.get(index);
        try {
            dropDownItem.click();
            System.out.println("==== click " + dropDownItem.getText() + " ====");
            sleep(200);
        } catch (Exception e) {

        }

        List<WebElement> dropdownChoices = new ArrayList<>();
        try {
            String controlId = dropDownItem.findElement(By.cssSelector("div[role=combobox]")).getAttribute("aria-controls");
            System.out.println("control id:" + controlId);
            dropdownChoices = webDriver.findElements(By.cssSelector("div[id*=\"" + controlId.substring(5) + "\"]" + " ul li"));
            System.out.println("选项:" + dropdownChoices.stream().map(e -> e.getText()).collect(Collectors.joining("-")));
        } catch (Exception e) {
            e.printStackTrace();
        }

        if (dropdownChoices.stream().anyMatch(e -> "2020".equals(e.getText())
            || "2021".equals(e.getText())
            || "2019".equals(e.getText())
            || "2018".equals(e.getText())
            || "2017".equals(e.getText()))) {
            // skip year
            System.out.println("skip year");
            return getAllTableData(webDriver, index + 1);
        }

        for (int i = 0; i < dropdownChoices.size(); i++) {
            WebElement dropDownChoice = dropdownChoices.get(i);
            try {
                sleep(50);
                dropDownItem.click();
                sleep(180);

                if (!"true".equals(dropDownChoice.getAttribute("aria-selected"))) {
                    try {
                        dropDownChoice.click();
                    } catch (Exception e) {

                    }
                    sleep(300);
                }
            } catch (Exception e) {
                e.printStackTrace();
            }

            checkLogin(webDriver);

            if (index < dropDownList.size() - 1) {
                List<List<String>> subDatas = getAllTableData(webDriver, index + 1);
                datas.addAll(subDatas);
            } else { // last page
                while (true) {
                    WebElement nextPage = null;
                    try {
                        nextPage = webDriver.findElement(By.cssSelector("#scoreline > div.province_score_line_table > div.table_pagination_box > div > div > ul > li.ant-pagination-next"));
                    } catch (Exception e) {
                        // no next page
                    }

                    List<String> dropDowns = webDriver.findElements(By.cssSelector("#scoreline > div.content-top.content_top_1_4 > div.scoreLine-dropDown > div.dropdown-box"))
                        .stream().map(WebElement::getText).collect(Collectors.toList());

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

                        dd.addAll(dropDowns);
                        for (int k = 0; k < tds.size(); k++) {
                            System.out.print(tds.get(k).getText());
                            System.out.print("\t");

                            try {
                                dd.add(tds.get(k).getText());
                            } catch (Exception e) {
                                dd.add("-");
                            }
                        }
                        System.out.println("");

                        if (dd.stream().allMatch(String::isEmpty)) {
                            // remove empty
                            continue;
                        }

                        datas.add(dd);
                    }

                    if (nextPage == null || nextPage.getAttribute("class").contains("ant-pagination-disabled")) {
                        break;
                    } else {
                        sleep(100);
                        nextPage.click();
                        sleep(300);
                    }
                }
            }
        }

        return datas;
    }

    static class Data {
        String school;
        List<List<String>> datas;
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

        for (int i = 0; i < data.datas.size(); i++) {
            if (data.datas.isEmpty()) {
                continue;
            }

            HSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
//                if (i == 0) {
//                    HSSFCellStyle style = workbook.createCellStyle();
//                    style.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
//                    row.setRowStyle(style);
//                }

            for (int j = 0; j < data.datas.get(i).size(); j++) {
                row.createCell(j).setCellValue(data.datas.get(i).get(j));
            }
        }

        HSSFRow endRow = sheet.createRow(sheet.getLastRowNum() + 1);
        endRow.createCell(0).setCellValue("");

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
