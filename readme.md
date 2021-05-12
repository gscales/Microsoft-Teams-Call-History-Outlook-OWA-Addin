# Microsoft Teams Call History Outlook/OWA Addin #

One of the Microsoft Teams features that I like the least is call history, for a number of reasons firstly its quite limited only 30 days worth of data and only includes direct calls and not meetings (which are often just scheduled calls). So when it comes to actually answering a question (like the one I found myself asking last week) when did I last have a call or meeting with person X it can't be used to answer this question. The other thing is having that information available in the context of Outlook/OWA is a lot more useful for me (anyway) then digging through anther client to find its doesn't answer the question anyway. However while the call history in Teams only goes back 30 days the Teams CDR (Call Data Records) which I covered in [https://dev.to/gscales/accessing-microsoft-teams-summary-records-cdr-s-for-calls-and-meetings-using-exchange-web-services-3581](https://www.blogger.com/u/1/blog/post/edit/7123450/8079661183382240070#) do go back as far as your mailbox folder retention period. So with a little work I can use Exchange Web Services to create an Outlook/OWA addin that can be used to show the Call history for a person inline in OWA when you active it on a message you received from that user Eg

![](https://1.bp.blogspot.com/-WJGIYKDZ2QM/YJjCw-AMckI/AAAAAAAAOGo/HcTdu6C4qH497My59jGYaJYe-xWpC2zoACLcBGAsYHQ/w640-h530/callhist1.PNG)

 

How this works is that the CDR item that is stored in the Mailbox stores the participants of the Call or Meeting in the recipients collection of the CDR mailbox item. The means you can make a KQL query on the participants to find all the CDR items that involve a particular user based on the sender address of the Email that the OWA Addin is activated on. eg the EWS Query ends up looking like

```
<soap:Envelope
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
	xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
	xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
	xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
	<soap:Header>
		<t:RequestServerVersion Version="Exchange2016" />
	</soap:Header>
	<soap:Body>
		<m:FindItem Traversal="Shallow">
			<m:ItemShape>
				<t:BaseShape>IdOnly</t:BaseShape>
				<t:AdditionalProperties>
					<t:FieldURI FieldURI="item:ItemClass"
						xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" />
						<t:ExtendedFieldURI PropertyTag="96" PropertyType="SystemTime" />
						<t:ExtendedFieldURI PropertyTag="97" PropertyType="SystemTime" />
						<t:ExtendedFieldURI PropertyTag="23809" PropertyType="String" />
					</t:AdditionalProperties>
				</m:ItemShape>
				<m:IndexedPageItemView MaxEntriesReturned="40" Offset="0" BasePoint="Beginning" />
				<m:ParentFolderIds>
					<t:FolderId Id="A=....." />
				</m:ParentFolderIds>
				<m:QueryString>participants:"e5tmp5@datarumble.com"</m:QueryString>
			</m:FindItem>
		</soap:Body>
	</soap:Envelope>
```

The extended properties that are used for the Start and End are pidTagStart and pidTagEnd which get set on the CDR mailbox item and then the SenderEmailSMTP address is used to work out who is the Organizer or the call initiator. To work out the type of Call the ItemClass is used so the last part of the ItemClass will tell you what the CRD is for eg call or meeting (If your wondering if you could do the same with the Microsoft Graph the answer is no because the Items in question are not a subclass of IPM.Note so aren't accessible when using the Microsoft Graph)

Want to give it a try yourself ?

I've hosted the files on my GitHub pages so its easy to test (if you like it clone it and host it somewhere else). But all you need to do is add it as a custom addin (if you allowed to) using the 
URL-

  https://gscales.github.io/TeamsCallHistory/TeamsCallHistory.xml

!![](https://1.bp.blogspot.com/-fV2Wxo0Cr7Q/XIHvqow9SVI/AAAAAAAACTM/Vo-Pc3Q74AgZ-_KNh_Nl9UhZmBbouJS1wCLcBGAs/s1600/addin3.JPG)
