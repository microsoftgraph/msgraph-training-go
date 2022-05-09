// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <GraphHelperSnippet>
package graphhelper

import (
	"context"
	"fmt"
	"os"
	"strings"

	"github.com/Azure/azure-sdk-for-go/sdk/azcore/policy"
	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
	auth "github.com/microsoft/kiota-authentication-azure-go"
	msgraphsdk "github.com/microsoftgraph/msgraph-sdk-go"
	"github.com/microsoftgraph/msgraph-sdk-go/me"
	"github.com/microsoftgraph/msgraph-sdk-go/me/mailfolders/item/messages"
	"github.com/microsoftgraph/msgraph-sdk-go/me/sendmail"
	"github.com/microsoftgraph/msgraph-sdk-go/models"
	"github.com/microsoftgraph/msgraph-sdk-go/users"
)

type GraphHelper struct {
	deviceCodeCredential   *azidentity.DeviceCodeCredential
	userClient             *msgraphsdk.GraphServiceClient
	graphUserScopes        []string
	clientSecretCredential *azidentity.ClientSecretCredential
	appClient              *msgraphsdk.GraphServiceClient
}

func NewGraphHelper() *GraphHelper {
	g := &GraphHelper{}
	return g
}

// </GraphHelperSnippet>

// <UserAuthConfigSnippet>
func (g *GraphHelper) InitializeGraphForUserAuth() error {
	clientId := os.Getenv("CLIENT_ID")
	authTenant := os.Getenv("AUTH_TENANT")
	scopes := os.Getenv("GRAPH_USER_SCOPES")
	g.graphUserScopes = strings.Split(scopes, ",")

	// Create the device code credential
	credential, err := azidentity.NewDeviceCodeCredential(&azidentity.DeviceCodeCredentialOptions{
		ClientID: clientId,
		TenantID: authTenant,
		UserPrompt: func(ctx context.Context, message azidentity.DeviceCodeMessage) error {
			fmt.Println(message.Message)
			return nil
		},
	})
	if err != nil {
		return err
	}

	g.deviceCodeCredential = credential

	// Create an auth provider using the credential
	authProvider, err := auth.NewAzureIdentityAuthenticationProviderWithScopes(credential, g.graphUserScopes)
	if err != nil {
		return err
	}

	// Create a request adapter using the auth provider
	adapter, err := msgraphsdk.NewGraphRequestAdapter(authProvider)
	if err != nil {
		return err
	}

	// Create a Graph client using request adapter
	client := msgraphsdk.NewGraphServiceClient(adapter)
	g.userClient = client

	return nil
}

// </UserAuthConfigSnippet>

// <GetUserTokenSnippet>
func (g *GraphHelper) GetUserToken() (*string, error) {
	token, err := g.deviceCodeCredential.GetToken(context.Background(), policy.TokenRequestOptions{
		Scopes: g.graphUserScopes,
	})
	if err != nil {
		return nil, err
	}

	return &token.Token, nil
}

// </GetUserTokenSnippet>

// <GetUserSnippet>
func (g *GraphHelper) GetUser() (models.Userable, error) {
	query := me.MeRequestBuilderGetQueryParameters{
		// Only request specific properties
		Select: []string{"displayName", "mail", "userPrincipalName"},
	}

	return g.userClient.Me().
		GetWithRequestConfigurationAndResponseHandler(
			&me.MeRequestBuilderGetRequestConfiguration{
				QueryParameters: &query,
			},
			nil)
}

// </GetUserSnippet>

// <GetInboxSnippet>
func (g *GraphHelper) GetInbox() (models.MessageCollectionResponseable, error) {
	var topValue int32 = 25
	query := messages.MessagesRequestBuilderGetQueryParameters{
		// Only request specific properties
		Select: []string{"from", "isRead", "receivedDateTime", "subject"},
		// Get at most 25 results
		Top: &topValue,
		// Sort by received time, newest first
		Orderby: []string{"receivedDateTime DESC"},
	}

	return g.userClient.Me().
		MailFoldersById("inbox").
		Messages().
		GetWithRequestConfigurationAndResponseHandler(
			&messages.MessagesRequestBuilderGetRequestConfiguration{
				QueryParameters: &query,
			},
			nil)
}

// </GetInboxSnippet>

// <SendMailSnippet>
func (g *GraphHelper) SendMail(subject *string, body *string, recipient *string) error {
	// Create a new message
	message := models.NewMessage()
	message.SetSubject(subject)

	messageBody := models.NewItemBody()
	messageBody.SetContent(body)
	contentType := models.TEXT_BODYTYPE
	messageBody.SetContentType(&contentType)
	message.SetBody(messageBody)

	toRecipient := models.NewRecipient()
	emailAddress := models.NewEmailAddress()
	emailAddress.SetAddress(recipient)
	toRecipient.SetEmailAddress(emailAddress)
	message.SetToRecipients([]models.Recipientable{
		toRecipient,
	})

	sendMailBody := sendmail.NewSendMailRequestBody()
	sendMailBody.SetMessage(message)

	// Send the message
	return g.userClient.Me().SendMail().Post(sendMailBody)
}

// </SendMailSnippet>

// <AppOnyAuthConfigSnippet>
func (g *GraphHelper) EnsureGraphForAppOnlyAuth() error {
	if g.clientSecretCredential == nil {
		clientId := os.Getenv("CLIENT_ID")
		tenantId := os.Getenv("TENANT_ID")
		clientSecret := os.Getenv("CLIENT_SECRET")
		credential, err := azidentity.NewClientSecretCredential(tenantId, clientId, clientSecret, nil)
		if err != nil {
			return err
		}

		g.clientSecretCredential = credential
	}

	if g.appClient == nil {
		// Create an auth provider using the credential
		authProvider, err := auth.NewAzureIdentityAuthenticationProviderWithScopes(g.clientSecretCredential, []string{
			"https://graph.microsoft.com/.default",
		})

		// Create a request adapter using the auth provider
		adapter, err := msgraphsdk.NewGraphRequestAdapter(authProvider)
		if err != nil {
			return err
		}

		// Create a Graph client using request adapter
		client := msgraphsdk.NewGraphServiceClient(adapter)
		g.appClient = client
	}

	return nil
}

// </AppOnyAuthConfigSnippet>

// <GetUsersSnippet>
func (g *GraphHelper) GetUsers() (models.UserCollectionResponseable, error) {
	err := g.EnsureGraphForAppOnlyAuth()
	if err != nil {
		return nil, err
	}

	var topValue int32 = 25
	query := users.UsersRequestBuilderGetQueryParameters{
		// Only request specific properties
		Select: []string{"displayName", "id", "mail"},
		// Get at most 25 results
		Top: &topValue,
		// Sort by display name
		Orderby: []string{"displayName"},
	}

	return g.appClient.Users().
		GetWithRequestConfigurationAndResponseHandler(
			&users.UsersRequestBuilderGetRequestConfiguration{
				QueryParameters: &query,
			},
			nil)
}

// </GetUsersSnippet>

// <MakeGraphCallSnippet>
func (g *GraphHelper) MakeGraphCall() error {
	// INSERT YOUR CODE HERE
	// Note: if using appClient, be sure to call EnsureGraphForAppOnlyAuth
	// before using it.
	// err := g.EnsureGraphForAppOnlyAuth()
	// if err != nil {
	// 	return nil, err
	// }

	return nil
}

// </MakeGraphCallSnippet>
