package com.wecool.poiexport;

import com.deepoove.poi.XWPFTemplate;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * @author bowafterrain [mazhaoming@vip.qq.com]
 * @date 2021-12-29 20:51
 */
public class WordExport {

    private final InputStream template;

    private Map<String, Object> model;

    private String blank = "    ";

    public WordExport(InputStream template) {
        this.template = template;
        model = new HashMap<>();
    }

    public WordExport(InputStream template, Map<String, Object> model) {
        this.template = template;
        this.model = model;
    }

    public Map<String, Object> getModel() {
        return model;
    }

    public void setModel(Map<String, Object> model) {
        this.model = model;
    }

    public String getBlank() {
        return blank;
    }

    public void setBlank(String blank) {
        this.blank = blank;
    }

    public Object put(String key, Object value) {
        return model.put(key, value);
    }

    public void export(OutputStream outputStream) throws IOException {
        this.replaceModelNullToBlank();
        XWPFTemplate.compile(template).render(model, outputStream).close();
    }

    private void replaceModelNullToBlank() {
        model.replaceAll((k, v) -> v == null ? blank : v);
    }
}
