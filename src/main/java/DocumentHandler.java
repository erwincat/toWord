import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

import java.io.*;
import java.util.HashMap;
import java.util.List;
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
// test.ftl为要装载的模板
                t = configuration.getTemplate("data.xml");
                t.setEncoding("utf-8");
            } catch (IOException e) {
                e.printStackTrace();
            }
            return t;
        }

        public Writer getWriter(String jobId){
// 输出文档路径及名称
            String savePath = "/Users/wuchangming/Documents/";
            File file = new File(savePath+"upload");
            if(!file.exists()){
                file.mkdirs();
            }
            File outFile = new File(savePath+"upload/"+jobId+".doc");
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


private void getData(Map dataMap) {

dataMap.put("name", "用户信息");
    dataMap.put("day", "用户信息");
    dataMap.put("bank", "用户信息");
    dataMap.put("bankAccount", "用户信息");
    dataMap.put("account", "用户信息");
    dataMap.put("phoneNUmber", "用户信息");
    dataMap.put("shouquanren", "用户信息");
    dataMap.put("year", "用户信息");
    dataMap.put("month", "用户信息");


}




    public void personnelImportWord() throws Exception{
        String path = "/Users/wuchangming/Documents/";
//        String jobId = getStringParameter("job_id");
        Map<String, Object> paraMap = new HashMap<String, Object>();
        Map dataMap = new HashMap();
//        paraMap.put("jobId", jobId);
        DocumentHandler dh = new DocumentHandler();
        Template t = dh.getTemplate();

        List<Map> list=ReadExcelTable.readTable();
        for (Map map : list){
            Writer out = dh.getWriter((String)map.get("name"));
            dh.createDoc(t,map, out);
        }

    }
//    public InputStream getInputStream() throws Exception
//    {
//        String path = "/Users/wuchangming/Documents/";
//        File file = new File(path+"upload");
//        if(!file.exists()){
//            file.mkdirs();
//        }
//        return new FileInputStream(path+"upload/"+fileName);
//    }
}
