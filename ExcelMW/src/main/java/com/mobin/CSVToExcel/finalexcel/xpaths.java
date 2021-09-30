package com.mobin.CSVToExcel.finalexcel;


import org.htmlcleaner.HtmlCleaner;
import org.htmlcleaner.TagNode;
import org.htmlcleaner.XPatherException;
import org.jsoup.Jsoup;

import java.io.IOException;

public class xpaths {

    public static void main(String[] args) throws IOException, XPatherException {

        String url = "https://mbd.baidu.com/newspage/data/landingsuper?context=%7B%22nid%22%3A%22news_9601717093207652821%22%7D&n_type=0&p_from=1";
        String contents = Jsoup.connect(url).post().html();

        HtmlCleaner hc = new HtmlCleaner();
        TagNode tn = hc.clean(contents);
        //代表class="article-title"的div标签下面的h2标签里面的内容
        String xpath = "//div[@class='article-title']/h2/text()";
        Object[] objects = tn.evaluateXPath(xpath);
        for (Object object : objects) {
            System.out.println(object);
        }
        System.out.println(objects.length);

        System.out.println("---------------------------------");
        //代表class="article-content"的div标签下面的p标签下的span标签里面的内容
        String xpath1 = "//div[@class='article-content']/p/span/text()";

        Object[] objects1 = tn.evaluateXPath(xpath1);
        for (Object object : objects1) {
            System.out.println(object);
        }

        System.out.println(objects1.length);
    }
}
