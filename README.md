# mail
java 中发送带有 excel 附件的邮件
先说下这个 springboot项目的背景,攻击需要监控金额变化的异常,通过 excel 附件的方式发送到领导人的邮箱里
如果对大家有用,请点个赞哦
文中采用的是字节流,直接将生成的 excel 输出流转为 inputstream,节省了中间的开发

首先上配置,我用的是 qq 邮箱作为邮箱服务器
  mail:
      username: xxxxx@qq.com
      password: kk**********ib
      host: smtp.qq.com
      receiver: XXXX
      cc: XXXXX,xxxxx@qq.com,xxxxxx,xxxxx


核心类是

package com.haitun.job.monitor.service.impl;

import com.haitun.job.monitor.support.EmailProperties;
import com.sun.mail.util.MailSSLSocketFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.annotation.Resource;
import javax.mail.*;
import javax.mail.internet.*;
import javax.mail.internet.MimeMessage.RecipientType;
import javax.mail.util.ByteArrayDataSource;
import java.io.ByteArrayInputStream;
import java.net.URLEncoder;
import java.security.GeneralSecurityException;
import java.util.Properties;

/**
 * author:May
 * Date: 2019/5/20.
 * Time: 上午10:49
 * 邮件的 service
 */
@Service
public class BaseMail {
    private static final Logger LOGGER = LoggerFactory.getLogger(BaseMail.class);

    @Resource
    private EmailProperties emailProperties;
    /**
     * 发送邮件
     * @param title     邮件标题
     * @param text      内容
     * @param text      附件标题
     * @param
     */
    public void sendMsgFileDs(String title, String text,String affixName, ByteArrayInputStream inputstream) {

        LOGGER.info("[邮箱配置参数]username是{},接受者是{},抄送者是{}",emailProperties.getUsername(),emailProperties.getReceiver(),emailProperties.getCc());
        Session session = assembleSession();
        Message msg = new MimeMessage(session);
        try {
            msg.setFrom(new InternetAddress(emailProperties.getUsername()));

            String encodedSubject = MimeUtility.encodeText(title, MimeUtility.mimeCharset("gb2312"), null);//防止乱码
            msg.setSubject(encodedSubject);
            msg.setRecipients(RecipientType.TO, acceptAddressList(emailProperties.getReceiver(),emailProperties.getCc()));
            MimeBodyPart contentPart = (MimeBodyPart) createContent(text, inputstream,affixName);//参数为正文内容和附件流
            MimeMultipart mime = new MimeMultipart("mixed");
            mime.addBodyPart(contentPart);
            msg.setContent(mime);
            Transport.send(msg);
        } catch (Exception e) {
            LOGGER.error("[邮件配置过程中异常]信息是{}",e.getMessage());
        }
    }

    public Address[] acceptAddressList(String acceptAddress,String cc) {
        // 创建邮件的接收者地址，并设置到邮件消息中
        Address[] tos = null;
        String [] copyEmail = cc.split(",");
        try {
            if(null != copyEmail){
                tos = new InternetAddress[copyEmail.length + 1];
                tos[0] = new InternetAddress(acceptAddress);
                for(int i=0;i<copyEmail.length;i++){
                    tos[i + 1] = new InternetAddress(copyEmail[i]);
                }
            } else {
                tos = new InternetAddress[1];
                tos[0] = new InternetAddress(acceptAddress);
            }
        } catch (AddressException e) {
            LOGGER.error("[邮件配置收件人异常]信息是{}",e.getMessage());
        }
        return tos;
    }

    public Session assembleSession() {
        Session session = null;
        Properties props = new Properties();
        props.setProperty("mail.smtp.auth", "true");
        props.setProperty("mail.transport.protocol", "smtp");
        props.setProperty("mail.smtp.port", "465");
        props.setProperty("mail.smtp.host", emailProperties.getHost());//邮件服务器
        //开启安全协议
        MailSSLSocketFactory sf = null;
        try {
            sf = new MailSSLSocketFactory();
            sf.setTrustAllHosts(true);
        } catch (GeneralSecurityException e1) {
            LOGGER.error("[邮件配置session异常]信息是{}",e1.getMessage());
        }
        props.put("mail.smtp.ssl.socketFactory", sf);
        props.put("mail.smtp.ssl.enable", "true");
        session = Session.getDefaultInstance(props, new MyAuthenricator(emailProperties.getUsername(), emailProperties.getPassword()));
        return session;
    }

    static Part createContent(String content, ByteArrayInputStream inputstream, String affixName) {
        MimeBodyPart contentPart = null;
        try {
            contentPart = new MimeBodyPart();
            MimeMultipart contentMultipart = new MimeMultipart("related");
            MimeBodyPart htmlPart = new MimeBodyPart();
            htmlPart.setContent(content, "text/html;charset=gbk");
            contentMultipart.addBodyPart(htmlPart);
            //附件部分
            MimeBodyPart excelBodyPart = new MimeBodyPart();
            DataSource dataSource = new ByteArrayDataSource(inputstream, "application/excel");
            DataHandler dataHandler = new DataHandler(dataSource);
            excelBodyPart.setDataHandler(dataHandler);
            excelBodyPart.setFileName(MimeUtility.encodeText(affixName));
            contentMultipart.addBodyPart(excelBodyPart);
            contentPart.setContent(contentMultipart);
        } catch (Exception e) {
            LOGGER.error("[邮件内容和附件配置异常]信息是{}",e.getMessage());
        }
        return contentPart;
    }

    //用户名密码验证，需要实现抽象类Authenticator的抽象方法PasswordAuthentication
    static class MyAuthenricator extends Authenticator {
        String u = null;
        String p = null;

        public MyAuthenricator(String u, String p) {
            this.u = u;
            this.p = p;
        }

        @Override
        protected PasswordAuthentication getPasswordAuthentication() {
            return new PasswordAuthentication(u, p);
        }
    }

}

然后是excel 处理类

package com.haitun.job.monitor.util;

import com.haitun.job.monitor.model.MonitorResponse;
import org.apache.http.client.utils.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Date;
import java.util.List;

/**
 * author:May
 * Date: 2019/5/21.
 * Time: 上午11:21
 */
public class EmailUtils {

    public static String emailContent2Html(List<MonitorResponse> monitorResponseList){

        StringBuilder content = new StringBuilder("<html><head></head><body><h2>Hi All:以下是资金异常的记录,\n附件中有完整的excle</h2>");
        content.append("<table border=\"5\" style=\"border:solid 1px #E8F2F9;font-size=14px;;font-size:18px;\">");
        content.append("<tr style=\"background-color: #428BCA; color:#ffffff\"><th>用户Id</th><th>用户名称</th><th>订单Id</th><th>异常原因</th></tr>");
        for (MonitorResponse data : monitorResponseList) {
            content.append("<tr>");
            content.append("<td>" + data.getUserId() + "</td>"); //第一列
            content.append("<td>" + data.getUsername() + "</td>"); //第二列
            content.append("<td>" + data.getOrderId() + "</td>"); //第三列
            content.append("<td>" + data.getReason() + "</td>"); //第三列
            content.append("</tr>");
        }
        content.append("</table>");
        content.append("<h3>马轶 May\n" +
                "后端开发工程师 \n" +
                "Mobile：18516145032\n" +
                "邮 箱：mayi@haitun.group\n" +
                "地 址：上海市嘉定区众仁路399号，运通星财富广场A座5楼</h3>");
        content.append("</body></html>");
        return content.toString();
    }
    /**
     * 发送邮件伴随着附件
     * @param monitorResponseList 充值异常参数
     */
    public static ByteArrayInputStream sendMailWithAttachment(List<MonitorResponse> monitorResponseList){

        String[] headers = {"用户Id","用户名称","订单Id","异常原因"};
        // 声明一个工作薄
        HSSFWorkbook wb = new HSSFWorkbook();
        // 生成一个表格
        HSSFSheet sheet = wb.createSheet();
        HSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            HSSFCell cell = row.createCell((short)i);
            cell.setCellValue(headers[i]);
        }
        int rowIndex = 1;
        for(int j=0; j<monitorResponseList.size(); j++){
            row = sheet.createRow(rowIndex);
            rowIndex++;
            HSSFCell cell1 = row.createCell((short)0);
            if(null != monitorResponseList.get(j)){
                cell1.setCellValue(monitorResponseList.get(j).getUserId());
            }
            cell1 = row.createCell((short)1);
            if(null != monitorResponseList.get(j)){
                cell1.setCellValue(monitorResponseList.get(j).getUsername());
            }
            cell1 = row.createCell((short)2);
            if(null != monitorResponseList.get(j)){
                cell1.setCellValue(monitorResponseList.get(j).getOrderId());
            }
            cell1 = row.createCell((short)3);
            if(null != monitorResponseList.get(j)){
                cell1.setCellValue(monitorResponseList.get(j).getReason());
            }
        }
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn((short)i);
        }

        ByteArrayOutputStream os = new ByteArrayOutputStream(1000);
        try{
            wb.write(os);
        }catch (Exception ex){
            ex.printStackTrace();
        }

        byte[] bytes = os.toByteArray();
        ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(bytes);
        try {
            os.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

       return  byteArrayInputStream;
    }




}

