package com.microsoft.ews;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.ComparisonMode;
import microsoft.exchange.webservices.data.core.enumeration.search.ContainmentMode;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class ReceiveEWSMail {
  private static ExchangeService service;

  private String fromStringTerm;

  private String subjectTerm;

  public static void main(String[] args) throws Exception {
    String fromStringTerm = "tao.zhong@hpe.com";
    String subjectTerm = "Signed Mail with 2 attachments.";

    ReceiveEWSMail receive = new ReceiveEWSMail();
    receive.initialize();

    receive.findItems(fromStringTerm, subjectTerm, 1);
  }

  public void initialize() throws URISyntaxException {
    String username = "cainzhong@cainzhong.win";
    String password = "Cisco01!";
    String uri = "https://outlook.office365.com/EWS/Exchange.asmx";
    this.fromStringTerm = "tao.zhong@hpe.com";
    this.subjectTerm = "Signed Mail with 2 attachments.";

    service = new ExchangeService();
    ExchangeCredentials credentials = new WebCredentials(username, password);
    service.setCredentials(credentials);
    service.setUrl(new URI(uri));

  }

  public void findItems(String fromStringTerm, String subjectTerm, int pageSize) throws Exception {
    ItemView view = new ItemView(pageSize);
    // view.setPropertySet(new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived));

    SearchFilter.ContainsSubstring fromTermFilter = new SearchFilter.ContainsSubstring(EmailMessageSchema.From, fromStringTerm);
    SearchFilter.ContainsSubstring subjectFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, subjectTerm, ContainmentMode.Substring, ComparisonMode.IgnoreCase);

    FindItemsResults<Item> findResults = service.findItems(WellKnownFolderName.Inbox, new SearchFilter.SearchFilterCollection(LogicalOperator.Or, fromTermFilter, subjectFilter), view);

    System.out.println("Total number of items found: " + findResults.getTotalCount());
    List msgDataList = new ArrayList();
    for (Item item : findResults) {
      Map messageData = new HashMap();
      messageData = this.readEmailItem(item.getId());
      System.out.println("subject : " + messageData.get("subject").toString());
      System.out.println("Sender : " + messageData.get("senderName").toString());
      msgDataList.add(messageData);
      // Do something with the item.
    }
  }

  /**
   * Reading one email at a time. Using Item ID of the email.
   * Creating a message data map as a return value.
   */
  public Map readEmailItem(ItemId itemId) {
    Map messageData = new HashMap();
    try {
      Item itm = Item.bind(service, itemId, PropertySet.FirstClassProperties);
      EmailMessage emailMessage = EmailMessage.bind(service, itm.getId());
      messageData.put("emailItemId", emailMessage.getId().toString());
      messageData.put("subject", emailMessage.getSubject().toString());
      messageData.put("fromAddress", emailMessage.getFrom().getAddress().toString());
      messageData.put("senderName", emailMessage.getSender().getName().toString());
      Date dateTimeCreated = emailMessage.getDateTimeCreated();
      messageData.put("SendDate", dateTimeCreated.toString());
      Date dateTimeRecieved = emailMessage.getDateTimeReceived();
      messageData.put("RecievedDate", dateTimeRecieved.toString());
      messageData.put("Size", emailMessage.getSize() + "");
      messageData.put("emailBody", emailMessage.getBody().toString());
    } catch (Exception e) {
      e.printStackTrace();
    }
    return messageData;

  }
}