package microsoftgraph

import (
	"context"

	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
	graphsdk "github.com/microsoftgraph/msgraph-sdk-go"
	graphmodels "github.com/microsoftgraph/msgraph-sdk-go/models"
	graphusers "github.com/microsoftgraph/msgraph-sdk-go/users"
	//other-imports
)

type MicrosoftGraphMailer struct {
	graphClient *graphsdk.GraphServiceClient
}

func NewMicrosoftGraphMailer(tenantId, clientId, clientSecret string) (*MicrosoftGraphMailer, error) {
	graphMailer := &MicrosoftGraphMailer{}
	// 使用MIME raw模式，直接传入smtp收到的原始邮件，不需要解析HTML/文本
	cred, _ := azidentity.NewClientSecretCredential(
		tenantId,
		clientId,
		clientSecret,
		nil,
	)

	graphClient, err := graphsdk.NewGraphServiceClientWithCredentials(
		cred, []string{"https://graph.microsoft.com/.default"})
	if err != nil {
		return nil, err
	}
	graphMailer.graphClient = graphClient
	return graphMailer, nil
}

func (p *MicrosoftGraphMailer) SendMail(bodyType, subject, from string, to []string, msg []byte) error {
	requestBody := graphusers.NewItemSendMailPostRequestBody()
	message := graphmodels.NewMessage()
	// subject
	message.SetSubject(&subject)
	// from
	recipient := graphmodels.NewRecipient()
	emailAddress := graphmodels.NewEmailAddress()
	emailAddress.SetAddress(&from)
	recipient.SetEmailAddress(emailAddress)
	message.SetFrom(recipient)
	// to
	toRecipients := []graphmodels.Recipientable{}
	for _, item := range to {
		recipient = graphmodels.NewRecipient()
		emailAddress = graphmodels.NewEmailAddress()
		emailAddress.SetAddress(&item)
		recipient.SetEmailAddress(emailAddress)
		toRecipients = append(toRecipients, recipient)
	}
	message.SetToRecipients(toRecipients)
	// body
	body := graphmodels.NewItemBody()
	contentType := graphmodels.TEXT_BODYTYPE
	if bodyType == "HTML" {
		contentType = graphmodels.HTML_BODYTYPE
	}
	body.SetContentType(&contentType)
	content := string(msg)
	body.SetContent(&content)
	//encoded := base64.StdEncoding.EncodeToString(msg)
	//body.SetContent(&encoded)

	message.SetBody(body)

	requestBody.SetMessage(message)
	saveToSentItems := true
	requestBody.SetSaveToSentItems(&saveToSentItems)
	return p.graphClient.Users().ByUserId(from).SendMail().Post(context.Background(), requestBody, nil)
}

func SendMail(from string, to []string, msg []byte) error {
	requestBody := graphusers.NewItemSendMailPostRequestBody()
	message := graphmodels.NewMessage()
	subject := "Meet for lunch?"
	message.SetSubject(&subject)
	body := graphmodels.NewItemBody()
	contentType := graphmodels.TEXT_BODYTYPE
	body.SetContentType(&contentType)
	content := "The new cafeteria is open."
	body.SetContent(&content)
	message.SetBody(body)

	recipient := graphmodels.NewRecipient()
	emailAddress := graphmodels.NewEmailAddress()
	address := "frannis@contoso.com"
	emailAddress.SetAddress(&address)
	recipient.SetEmailAddress(emailAddress)

	toRecipients := []graphmodels.Recipientable{
		recipient,
	}
	message.SetToRecipients(toRecipients)

	requestBody.SetMessage(message)
	saveToSentItems := true
	requestBody.SetSaveToSentItems(&saveToSentItems)

	cred, _ := azidentity.NewClientSecretCredential(
		"TENANT_ID",
		"CLIENT_ID",
		"CLIENT_SECRET",
		nil,
	)

	graphClient, _ := graphsdk.NewGraphServiceClientWithCredentials(
		cred, []string{"https://graph.microsoft.com/.default"})

	// To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=go
	graphClient.Users().ByUserId(from).SendMail().Post(context.Background(), requestBody, nil)
	return nil
}
