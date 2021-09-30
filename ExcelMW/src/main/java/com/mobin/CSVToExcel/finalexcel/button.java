package com.mobin.CSVToExcel.finalexcel;
import com.gargoylesoftware.htmlunit.WebClient;
import com.gargoylesoftware.htmlunit.html.HtmlElement;
import com.gargoylesoftware.htmlunit.html.HtmlInput;
import com.gargoylesoftware.htmlunit.html.HtmlPage;

import java.util.List;
public class button {




    /**
     * 模拟点击，动态获取页面信息
     * @author linhongcun
     *
     */

        public static void main(String[] args) throws Exception {
            // 创建webclient
            WebClient webClient = new WebClient();
            // 取消 JS 支持
            webClient.getOptions().setJavaScriptEnabled(false);
            // 取消 CSS 支持
            webClient.getOptions().setCssEnabled(false);
            // 获取指定网页实体
            HtmlPage page = (HtmlPage) webClient.getPage("https://bi.sankuai.com/dashboard/20858");
            // 获取搜索输入框
            //HtmlInput input = (HtmlInput) page.getHtmlElementById("input");
            // 往输入框 “填值”
            //input.setValueAttribute("larger5");
            // 获取搜索按钮
            //HtmlInput btn = (HtmlInput) page.getHtmlElementById("search-button");
            //Thread.sleep(5000);
            HtmlInput btn1 = (HtmlInput) page.getFirstByXPath("/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[2]/div/div/button[1]/span/span");
            //HtmlInput btn12 = (HtmlInput) page.get("/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[2]/div/div/button[1]/span/span");
            //System.out.println(btn12);
            // “点击” 搜索
            HtmlPage page2 = btn1.click();
            // 选择元素
            List<HtmlElement> spanList=page2.getByXPath("//h3[@class='res-title']/a");
            for(int i=0;i<spanList.size();i++) {
                // 输出新页面的文本
                System.out.println(i+1+"、"+spanList.get(i).asText());
            }
        }



}
