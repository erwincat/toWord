import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

import java.io.*;
import java.util.Map;

/**
 * Created by wuchangming on 17/1/6.
 */
public class DocumentHandler {


        private Configuration configuration = null;

        public DocumentHandler() {
            configuration = new Configuration();
            configuration.setDefaultEncoding("utf-8");
        }

        public Template getTemplate() throws IOException {
            configuration.setDirectoryForTemplateLoading(new File("/Users/wuchangming/Documents"));
            Template t = null;
            try {
                // data.xml为要装载的模板
                t = configuration.getTemplate("data.xml");
                t.setEncoding("utf-8");
            } catch (IOException e) {
                e.printStackTrace();
            }
            return t;
        }

        public Writer getWriter(String wordName){
            String savePath = "/Users/wuchangming/Documents/";
            File file = new File(savePath+"upload");
            if(!file.exists()){
                file.mkdirs();
            }
            File outFile = new File(savePath+"upload/"+wordName+".doc");
            Writer out = null;
            try {
                out = new BufferedWriter(new OutputStreamWriter(
                        new FileOutputStream(outFile), "utf-8"));
            } catch (Exception e1) {
                e1.printStackTrace();
            }
            return out;
        }


        public void createDoc(Template t,Map dataMap,Writer out) {
            try {
                t.process(dataMap, out);
                out.close();
            } catch (TemplateException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }


}
