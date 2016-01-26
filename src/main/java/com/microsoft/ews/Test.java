package com.microsoft.ews;


import java.net.URI;

import org.apache.commons.codec.binary.BinaryCodec;

import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.MessageBody;


public class Test {
  static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
    public boolean autodiscoverRedirectionUrlValidationCallback(String redirectionUrl) {
      return redirectionUrl.toLowerCase().startsWith("https://");
    }
  }

  public static void main(String[] args) throws Exception {
    String username = "cainzhong@cainzhong.win";
    String password = "Cisco01!";
    String uri = "https://outlook.office365.com/EWS/Exchange.asmx";

    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
    ExchangeCredentials credentials = new WebCredentials(username, password);
    service.setCredentials(credentials);
    // service.setUrl(new URI("https://" + "邮箱服务器地址" + "/EWS/Exchange.asmx"));
    // service.autodiscoverUrl(username, new RedirectionUrlCallback());
    service.setUrl(new URI(uri));

    System.out.println(BinaryCodec.class.getProtectionDomain().getCodeSource().getLocation());

    EmailMessage msg = new EmailMessage(service);
    msg.setSubject("Hello world!");
    msg.setBody(MessageBody.getMessageBodyFromText("Sent using the EWS Java API."));
    msg.getToRecipients().add(username);
    System.out.println(BinaryCodec.class.getProtectionDomain().getCodeSource().getLocation());

    msg.send();
    System.out.println(BinaryCodec.class.getProtectionDomain().getCodeSource().getLocation());

    System.out.println("Success!!!");
    // // Bind to the Inbox.
    // Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox);
    // System.out.println(inbox.getDisplayName());
    // ItemView view = new ItemView(10);
    // FindItemsResults<Item> findResults = service.findItems(inbox.getId(), view);
    // for (Item item : findResults.getItems()) {
    // EmailMessage message = EmailMessage.bind(service, item.getId());
    // System.out.println(message.getSender());
    // System.out.println("Sub -->" + item.getSubject());
    // }

  }
}
