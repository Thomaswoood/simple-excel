package com.thomas.alib.excel.loader;

import com.thomas.alib.excel.utils.StringUtils;
import org.apache.poi.util.IOUtils;

import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;

/**
 * 默认图片加载器，认为图片数据源是网络上完整的url来加载
 */
public class PictureLoaderDefault implements PictureLoader<String> {

    /**
     * 默认加载图片为byte数组方法，认为图片数据源是网络上完整的url来加载
     *
     * @param s 图片url
     * @return 图片byte数组
     * @throws IOException 图片读取IO错误
     */
    @Override
    public byte[] loadPicture(String s) throws IOException {
        if (StringUtils.isNotBlank(s) && !StringUtils.equals(s, "null")) {
            return null;
        }
        //加载网络图片到
        URL url = new URL(s);
        URLConnection conn = url.openConnection();
        try (InputStream inputStream = conn.getInputStream()) {
            return IOUtils.toByteArray(inputStream);
        }
    }
}
