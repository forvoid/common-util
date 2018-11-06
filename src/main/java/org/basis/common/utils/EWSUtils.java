package org.basis.common.utils;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.ComparisonMode;
import microsoft.exchange.webservices.data.core.enumeration.search.ContainmentMode;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.AttachmentCollection;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;
import org.apache.commons.lang.StringUtils;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;

/**
 * 微软发邮件
 *
 * @author taowenxiang
 * @date 2018/10/31
 * @since 1.0
 */
public class EWSUtils {

    /**
     * 邮箱账号
     */
    private static final String USER_NAME = "example@mail.com";

    /**
     * 邮箱密码
     */
    private static final String PASSWORD = "passwod";

    /**
     * 邮箱服务路径
     */
    private static final String MAIL_SERVICES_URL = "https://mail.example.com/EWS/Exchange.asmx";

    /**
     * 获取指定名称邮件信息的附件数据
     */
    public static String getEmailMessageAttachment(String emailName,String attachmentName) throws Exception {
        String result = null;
        // 获取邮件
        EmailMessage message = getEmailMessageByName(emailName);
        if (message == null) {
            // 如果邮件没有
            return null;
        }

        AttachmentCollection attachments = message.getAttachments();
        for (Attachment attachment : attachments) {
            if (StringUtils.equals(attachment.getName(),attachmentName)) {
                FileAttachment fat = (FileAttachment)attachment;
                System.out.println(fat.getContentType());
                fat.load("/tmp/" + fat.getName());
                result = "/tmp/" + fat.getName();
            }
        }
        return result;
    }

    /**
     * 根据名称查找对应的邮件信息 模糊查询
     *
     * @param name 模糊的名字
     * @return message 可能返回 null
     */
    public static EmailMessage getEmailMessageByName(String name) {
        ExchangeService service = getExchangeService();
        //绑定收件箱,同样可以绑定发件箱
        EmailMessage message = null;
        try {
            Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox);
            SearchFilter.ContainsSubstring subjectFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject,
                    name, ContainmentMode.Substring, ComparisonMode.IgnoreCase);
            ItemView view = new ItemView(1);
            FindItemsResults<Item> findResults = service.findItems(inbox.getId(), subjectFilter, view);
            for (Item item : findResults) {
                message = EmailMessage.bind(service, item.getId());
            }
        } catch (Throwable e) {
            // ignore
        }
        return message;
    }

    /**
     * 收取邮件
     *
     * @param max 最大收取邮件数
     * @throws Exception 报错
     */
    public static ArrayList<EmailMessage> receive(int max) throws Exception {
        ExchangeService service = getExchangeService();
        //绑定收件箱,同样可以绑定发件箱
        Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox);
        //获取文件总数量
        int count = inbox.getTotalCount();
        if (max > 0) {
            count = count > max ? max : count;
        }
        //循环获取邮箱邮件
        ItemView view = new ItemView(count);
        FindItemsResults<Item> findResults = service.findItems(inbox.getId(), view);
        ArrayList<EmailMessage> result = new ArrayList<>();
        for (Item item : findResults.getItems()) {
            EmailMessage message = EmailMessage.bind(service, item.getId());
            result.add(message);
        }
        return result;
    }

    /**
     * 发送带附件的mail
     *
     * @param subject         邮件标题
     * @param to              收件人列表
     * @param cc              抄送人列表
     * @param bodyText        邮件内容
     * @param attachmentPaths 附件地址列表
     * @throws Exception 报错信息
     */
    public static void send(String subject, String[] to, String[] cc, String bodyText,
                            String[] attachmentPaths,String attachmentContxtType)
            throws Exception {
        ExchangeService service = getExchangeService();

        EmailMessage msg = new EmailMessage(service);
        msg.setSubject(subject);
        MessageBody body = MessageBody.getMessageBodyFromText(bodyText);
        body.setBodyType(BodyType.HTML);
        msg.setBody(body);
        for (String toPerson : to) {
            msg.getToRecipients().add(toPerson);
        }
        if (cc != null) {
            for (String ccPerson : cc) {
                msg.getCcRecipients().add(ccPerson);
            }
        }
        if (attachmentPaths != null) {
            for (String attachmentPath : attachmentPaths) {
                FileAttachment fileAttachment = msg.getAttachments().addFileAttachment(attachmentPath);
                fileAttachment.setContentType(attachmentContxtType);
            }
        }
        msg.send();
    }

    /**
     * 创建邮件服务
     *
     * @return 邮件服务
     */
    private static ExchangeService getExchangeService() {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
        //用户认证信息
        ExchangeCredentials credentials;
        credentials = new WebCredentials(USER_NAME, PASSWORD);
        service.setCredentials(credentials);
        try {
            service.setUrl(new URI(MAIL_SERVICES_URL));
        } catch (URISyntaxException e) {
            e.printStackTrace();
        }
        return service;
    }


}
