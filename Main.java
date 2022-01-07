package com.example.testcallchrome;

import com.codeborne.selenide.SelenideElement;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Cookie;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static com.codeborne.selenide.Selenide.$;
import static com.codeborne.selenide.Selenide.$x;
import static com.codeborne.selenide.Selenide.sleep;

// https://www.jetbrains.com/
public class MainPage {
    public SelenideElement seeAllToolsButton = $("a.wt-button_mode_primary");
    public SelenideElement toolsMenu = $x("//div[contains(@class, 'menu-main__item') and text() = 'Developer Tools']");
    public SelenideElement searchButton = $("[data-test='menu-main-icon-search']");


    public static void main(String[] args) {
        System.setProperty("webdriver.chrome.driver", "C:/Users/yang/Downloads/chromedriver_win32/chromedriver.exe");

        WebDriver webDriver = new ChromeDriver();

        webDriver.get("https://gkcx.eol.cn/login");

        Cookie cookie = new Cookie.Builder("originLoginKey", "0ce562890884021dea25f8bf8affcca98bc64e850d19f13a9a3ea0984d142743dd661c5931ceab814df7e68079ab75ddf0b855843ead150c4b44567e52ae30870010e6531ff8cbf40552cece8d687092a00eb46c847953c9f85f5bcad7b878d606bba49ecbfdc2bdd96c4de8e7c91d8a0734087300b8499a66e0cfc1e6c5724733f5a5a0adf05339b2a6a44b18a77c82")
            .domain(".eol.cn")
            .path("/")
            .build();

        webDriver.manage().addCookie(cookie);

        cookie = new Cookie.Builder("parseLoginKey", "{%22login_type%22:%221%22%2C%22phone%22:%2216602900618%22%2C%22wx_uin%22:%22%22%2C%22mac%22:%22838a156daa230756a86bac3797313319%22%2C%22agent%22:%226%22%2C%22time%22:%221641305824%22%2C%22is_perfect%22:%221%22%2C%22random%22:%22XNjzXPwA%22%2C%22pushcode%22:%22%22}")
            .domain(".eol.cn")
            .path("/")
            .build();
        webDriver.manage().addCookie(cookie);

        webDriver.get("https://gkcx.eol.cn/login");

        WebElement loginBtn = webDriver.findElement(By.cssSelector("#root > div > div.container > div > div > div > div.login_set > div.login_wrap > div > div.inp_wrap > div.login_btn.cursor"));
        if (loginBtn != null) {
            // login
            while (true) {
                sleep(500);
                String currentUrl = webDriver.getCurrentUrl();
                System.out.println(currentUrl);
                if ("https://gkcx.eol.cn/".equalsIgnoreCase(currentUrl)) {
                    break;
                }
            }
        }

        startProcess(webDriver);




//        for (let i = 1 ;i <= childCount ; i++) {
//            let child =  document.querySelector("#b5776711-83dc-4f80-84ad-171e94a4f5bf > ul > li:nth-child("+i+")");
//            child.click();
//            let trs = document.querySelectorAll("#scoreline > div.province_score_line_table > div.line_table_box.major_score_table > table > tbody > tr");  let trCount = trs.length;
//            for(let j = 1 ;j < trCount; j++) {
//                let tr = trs[j]; let tds = tr.querySelectorAll("td");
//                console.log(tds[0].innerText, "\t",  tds[1].innerText, '\t', tds[2].innerText, "\t", tds[3].innerText); }
//        }


    }

    static void startProcess(WebDriver webDriver) {
        webDriver.get("https://gkcx.eol.cn/school/140/provinceline");
        String title = webDriver.getTitle();
        //let childCount = document.querySelector("#b5776711-83dc-4f80-84ad-171e94a4f5bf > ul").childElementCount;

        // #scoreline > div.content-top.content_top_1_4 > div > div:nth-child(1) > div > div > div > div
        WebElement element = webDriver.findElement(By.cssSelector("#scoreline > div.content-top.content_top_1_4 > div > div:nth-child(1) > div > div > div > div"));
        element.click();
        System.out.println(element.getText());

        sleep(50);

//        WebElement element2 = webDriver.findElement(By.cssSelector("#scoreline > div.content-top.content_top_1_4 > div > div:nth-child(2)"));
//        element2.click();
//        System.out.println(element2.getText());

        List<WebElement> lis = webDriver.findElements(By.cssSelector("ul[role=listbox] > li"));

        Data data = new Data();
        data.school = "清华大学";

        Map<String, List<List<String>>> datas = new LinkedHashMap<>();
        data.datas = datas;

        for (int i = 0; i < lis.size(); ) {
            try {
                WebElement child = lis.get(i);
                element.click();
                sleep(100);
                child.click();
                sleep(200);

                System.out.println("=============" + element.getText() + "=============");
                List<List<String>> d = new ArrayList<>();
                datas.put(element.getText(), d);

                while (true) {
                    WebElement nextPage = webDriver.findElement(By.cssSelector("#scoreline > div.province_score_line_table > div.table_pagination_box > div > div > ul > li.ant-pagination-next"));
                    List<WebElement> trs = webDriver.findElements(By.cssSelector("#scoreline > div.province_score_line_table > div.line_table_box.major_score_table > table > tbody > tr"));
                    for (int j = 0; j < trs.size(); j++) {
                        List<String> dd = new ArrayList<>();
                        d.add(dd);

                        WebElement tr = trs.get(j);
                        List<WebElement> tds = tr.findElements(By.cssSelector("td"));
                        for (int k = 0; k < tds.size(); k++) {
                            System.out.print(tds.get(k).getText());
                            System.out.print("\t");
                            dd.add(tds.get(k).getText());
                        }
                        System.out.println("");
                    }

                    if (nextPage == null || nextPage.getAttribute("class").contains("ant-pagination-disabled")) {
                        break;
                    } else {
                        nextPage.click();
                        sleep(100);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
                continue;
            }

            i++;
        }

        createReport(data);

        webDriver.close();
    }

    static class Data {
        String school;
        Map<String, List<List<String>>> datas;



    }

    private static File createReport(Data data) {
        File file = null;
        try {
            file = File.createTempFile(data.school, ".xls");  //临时文件
            System.out.println(file.getAbsolutePath());
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }

        HSSFWorkbook workbook = new HSSFWorkbook();
        data.datas.forEach((province, value) -> {

            HSSFSheet sheet = workbook.createSheet(province);

            HSSFRow rowhead = sheet.createRow((short) 0);

            //专业名称	录取批次	平均分	最低分/最低位次
            rowhead.createCell(0).setCellValue("专业名称");
            rowhead.createCell(1).setCellValue("录取批次");
            rowhead.createCell(2).setCellValue("平均分");
            rowhead.createCell(3).setCellValue("最低分/最低位次");

            List<List<String>> datas = value;
            for (int i = 0; i < datas.size(); i++) {
                HSSFRow row = sheet.createRow((i + 1));

                List<String> d = datas.get(i);
                for (int j = 0; j < d.size(); j ++) {
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
