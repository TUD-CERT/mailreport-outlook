<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0" xsi:type="MailApp">
  <Id>__ID__</Id>
  <Version>__VERSION__</Version>
  <ProviderName>__PROVIDER_NAME__</ProviderName>
  <DefaultLocale>en</DefaultLocale>
  <DisplayName DefaultValue="__MSG_extensionName_en__">
    <Override Locale="de" Value="__MSG_extensionName_de__"/>
  </DisplayName>
  <Description DefaultValue="__MSG_extensionDescription_en__">
    <Override Locale="de" Value="__MSG_extensionDescription_de__"/>
  </Description>
  <IconUrl DefaultValue="__HOSTED_AT__/assets/app_64.png"/>
  <HighResolutionIconUrl DefaultValue="__HOSTED_AT__/assets/app_128.png"/>
  <AppDomains>
    <AppDomain>__HOSTED_AT__</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="__HOSTED_AT__/report_fraud.html"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteMailbox</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadGroup">
                <Label resid="GroupLabel"/>
                <Control xsi:type="Menu" id="MenuButton">
                  <Label resid="MenuButton.Label"/>
                  <Supertip>
                    <Title resid="MenuButton.Label"/>
                    <Description resid="MenuButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Items>
                    <Item id="FraudReport">
                      <Label resid="FraudReport.Label"/>
                      <Supertip>
                        <Title resid="FraudReport.Label"/>
                        <Description resid="FraudReport.Tooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Fraud.16x16"/>
                        <bt:Image size="32" resid="Fraud.32x32"/>
                        <bt:Image size="80" resid="Fraud.80x80"/>
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <SourceLocation resid="ReportFraud.Url"/>
                      </Action>
                    </Item>
                    <Item id="SpamReport">
                      <Label resid="SpamReport.Label"/>
                      <Supertip>
                        <Title resid="SpamReport.Label"/>
                        <Description resid="SpamReport.Tooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Spam.16x16"/>
                        <bt:Image size="32" resid="Spam.32x32"/>
                        <bt:Image size="80" resid="Spam.80x80"/>
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>reportSpam</FunctionName>
                      </Action>
                    </Item>
                    <Item id="Options">
                      <Label resid="Options.Label"/>
                      <Supertip>
                        <Title resid="Options.Label"/>
                        <Description resid="Options.Tooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Options.16x16"/>
                        <bt:Image size="32" resid="Options.32x32"/>
                        <bt:Image size="80" resid="Options.80x80"/>
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <SourceLocation resid="Options.Url"/>
                      </Action>
                    </Item>
                  </Items>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="__HOSTED_AT__/assets/app_16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="__HOSTED_AT__/assets/app_32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="__HOSTED_AT__/assets/app_80.png"/>
        <bt:Image id="Fraud.16x16" DefaultValue="__HOSTED_AT__/assets/fraud_16.png"/>
        <bt:Image id="Fraud.32x32" DefaultValue="__HOSTED_AT__/assets/fraud_32.png"/>
        <bt:Image id="Fraud.80x80" DefaultValue="__HOSTED_AT__/assets/fraud_80.png"/>
        <bt:Image id="Spam.16x16" DefaultValue="__HOSTED_AT__/assets/spam_16.png"/>
        <bt:Image id="Spam.32x32" DefaultValue="__HOSTED_AT__/assets/spam_32.png"/>
        <bt:Image id="Spam.80x80" DefaultValue="__HOSTED_AT__/assets/spam_80.png"/>
        <bt:Image id="Options.16x16" DefaultValue="__HOSTED_AT__/assets/options_16.png"/>
        <bt:Image id="Options.32x32" DefaultValue="__HOSTED_AT__/assets/options_32.png"/>
        <bt:Image id="Options.80x80" DefaultValue="__HOSTED_AT__/assets/options_80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="__HOSTED_AT__/commands.html"/>
        <bt:Url id="ReportFraud.Url" DefaultValue="__HOSTED_AT__/report_fraud.html"/>
        <bt:Url id="Options.Url" DefaultValue="__HOSTED_AT__/options.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="__MSG_extensionGroup_en__">
          <bt:Override Locale="de" Value="__MSG_extensionGroup_de__"/>
        </bt:String>
        <bt:String id="MenuButton.Label" DefaultValue="__MSG_extensionButtonLabel_en__">
          <bt:Override Locale="de" Value="__MSG_extensionButtonLabel_de__"/>
        </bt:String>
        <bt:String id="FraudReport.Label" DefaultValue="__MSG_extensionFraudReportLabel_en__">
          <bt:Override Locale="de" Value="__MSG_extensionFraudReportLabel_de__"/>
        </bt:String>
        <bt:String id="SpamReport.Label" DefaultValue="__MSG_extensionSpamReportLabel_en__">
          <bt:Override Locale="de" Value="__MSG_extensionSpamReportLabel_de__"/>
        </bt:String>
        <bt:String id="Options.Label" DefaultValue="__MSG_extensionOptionsLabel_en__">
          <bt:Override Locale="de" Value="__MSG_extensionOptionsLabel_de__"/>
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="MenuButton.Tooltip" DefaultValue="__MSG_extensionButtonTooltip_en__">
          <bt:Override Locale="de" Value="__MSG_extensionButtonTooltip_de__"/>
        </bt:String>
        <bt:String id="FraudReport.Tooltip" DefaultValue="__MSG_extensionFraudReportTooltip_en__">
          <bt:Override Locale="de" Value="__MSG_extensionFraudReportTooltip_de__"/>
        </bt:String>
        <bt:String id="SpamReport.Tooltip" DefaultValue="__MSG_extensionSpamReportTooltip_en__">
          <bt:Override Locale="de" Value="__MSG_extensionSpamReportTooltip_de__"/>
        </bt:String>
        <bt:String id="Options.Tooltip" DefaultValue="__MSG_extensionOptionsTooltip_en__">
          <bt:Override Locale="de" Value="__MSG_extensionOptionsTooltip_de__"/>
        </bt:String>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>