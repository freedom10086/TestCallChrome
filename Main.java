package com.example.testcallchrome;

import org.apache.commons.lang3.StringUtils;
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

import static com.codeborne.selenide.Selenide.sleep;

public class MainPage {

    public static void main(String[] args) {
        System.setProperty("webdriver.chrome.driver", "chromedriver.exe");

        ChromeOptions chromeOptions =  new ChromeOptions();
        chromeOptions.addArguments("user-data-dir=C:/Users/yang/AppData/Local/Google/Chrome/User Data/Default");

        WebDriver webDriver = new ChromeDriver(chromeOptions);

        webDriver.get("https://www.gaokao.cn/school/search");

        while (true) {
            WebElement loginBtn = null;
            try {
                loginBtn = webDriver.findElement(By.cssSelector("div.nologin.float_r.cancle_select_bg"));
            } catch (Exception e) {
                e.printStackTrace();
            }

            if (loginBtn != null) {
                try {
                    loginBtn.click();
                } catch (Exception e) {
                }
                sleep(5000);
            } else {
                // start process
                // get school list

                while (true) {
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
                        processCnt ++;
                        break;
                    }

                    if (!"https://www.gaokao.cn/school/search".equals(webDriver.getCurrentUrl())) {
                        webDriver.get("https://www.gaokao.cn/school/search");
                        sleep(100);
                    }

                    WebElement nextPageBtn = null;
                    try {
                        nextPageBtn = webDriver.findElement(By.cssSelector("#root i.anticon.anticon-right"));
                    } catch (Exception e) {

                    }

                    if (nextPageBtn != null) {
                        if (processCnt == 0) {
                            nextPageBtn.click();
                            processCnt = 0;
                            sleep(120);
                        }
                    } else {
                        break;
                    }
                }

                webDriver.close();
            }
        }


    }

    static void startProcess(WebDriver webDriver, String schoolName, String schoolIndex) {
        webDriver.get("https://gkcx.eol.cn/school/" + schoolIndex + "/provinceline");
        sleep(200);
        String title = webDriver.getTitle();
        //let childCount = document.querySelector("#b5776711-83dc-4f80-84ad-171e94a4f5bf > ul").childElementCount;

        // #scoreline > div.content-top.content_top_1_4 > div > div:nth-child(1) > div > div > div > div
        WebElement provinceElement = webDriver.findElement(By.cssSelector("#scoreline > div.content-top.content_top_1_4 > div > div:nth-child(1) > div > div > div > div"));
        provinceElement.click();
        sleep(120);

//        WebElement element2 = webDriver.findElement(By.cssSelector("#scoreline > div.content-top.content_top_1_4 > div > div:nth-child(2)"));
//        element2.click();
//        System.out.println(element2.getText());

        List<WebElement> provinceList = webDriver.findElements(By.cssSelector("ul[role=listbox] > li"));

        Data data = new Data();
        data.school = schoolName;

        Map<String, List<List<String>>> datas = new LinkedHashMap<>();
        data.datas = datas;

        for (int i = 0; i < provinceList.size(); ) {
            try {
                WebElement province = provinceList.get(i);
                provinceElement.click();
                sleep(120);
                province.click();
                sleep(200);

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
                        List<WebElement> tds = tr.findElements(By.cssSelector("td"));
                        for (int k = 0; k < tds.size(); k++) {
                            System.out.print(tds.get(k).getText());
                            System.out.print("\t");
                            dd.add(tds.get(k).getText());
                        }
                        System.out.println("");

                        if (dd.stream().allMatch(StringUtils::isEmpty)) {
                            // remove empty
                            continue;
                        }

                        d.add(dd);
                    }

                    if (nextPage == null || nextPage.getAttribute("class").contains("ant-pagination-disabled")) {
                        break;
                    } else {
                        nextPage.click();
                        sleep(120);
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
        data.datas.forEach((province, datas) -> {
            HSSFSheet sheet = workbook.createSheet(province);
            for (int i = 0; i < datas.size(); i++) {
                List<String> d = datas.get(i);
                if (d.isEmpty()) {
                    continue;
                }
                HSSFRow row = sheet.createRow(i);
                if (i == 0) {
                    HSSFCellStyle style = workbook.createCellStyle();
                    style.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
                    row.setRowStyle(style);
                }

                for (int j = 0; j < d.size(); j++) {
                    row.createCell(j).setCellValue(d.get(j));
                }
            }
        });

//        //auto column width 自适应列宽
//        HSSFRow row = workbook.getSheetAt(0).getRow(0);
//        for (int colNum = 0; colNum < row.getLastCellNum(); colNum++) {
//            workbook.getSheetAt(0).autoSizeColumn(colNum);
//        }

        try {
            FileOutputStream fileOut = new FileOutputStream(file);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return file;
    }
}
