﻿<?xml version="1.0" encoding="UTF-8"?>
<!--Created:ce44715c-8c4e-446b-879c-ea9ebe0f09c8-->
<OfficeApp 
          xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
          xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" 
          xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0" 
          xsi:type="MailApp">

  <!-- Begin Basic Settings: Add-in metadata, used for all versions of Office unless override provided. -->

  <!-- IMPORTANT! Id must be unique for your add-in, if you reuse this manifest ensure that you change this id to a new GUID. -->
  <Id>b1548094-48cf-4a92-a27c-da7837265cbb</Id>

  <!--Version. Updates from the store only get triggered if there is a version change. -->
  <Version>1.0.0.0</Version>
  <ProviderName>Pineland</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
  <DisplayName DefaultValue="TicketSystem" />
  <Description DefaultValue="TicketSystem"/>
  <IconUrl DefaultValue="https://imgur.com/PSSESTN"/>

  <SupportUrl DefaultVal![image](https://github.com/jcanady6/TicketSystem/assets/43898717/cb2980e9-215f-4fac-82c3-397c9a60e393)
ue="http://www.contoso.com" />
  <!-- Domains that will be allowed when navigating. For example, if you use ShowTaskpane and then have an href link, navigation will only be allowed if the domain is on this list. -->
  <AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
    <AppDomain>AppDomain3</AppDomain>
  </AppDomains>
  <!--End Basic Settings. -->
  
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://github.com/jcanady6/TicketSystem/blob/master/TicketSystemWeb/MessageRead.html"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">

        <DesktopFormFactor>
          <!-- Location of the Functions that UI-less buttons can trigger (ExecuteFunction Actions). -->
          <FunctionFile resid="functionFile" />

          <!-- Message Read -->
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Use the default tab of the ExtensionPoint or create your own with <CustomTab id="myTab"> -->
            <OfficeTab id="TabDefault">
              <!-- Up to 6 Groups added per Tab -->
              <Group id="msgReadGroup">
                <Label resid="groupLabel" />
                <!-- Launch the add-in : task pane button -->
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="paneReadButtonLabel" />
                  <Supertip>
                    <Title resid="paneReadSuperTipTitle" />
                    <Description resid="paneReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="TicketRed16" />
                    <bt:Image size="32" resid="TicketRed32" />
                    <bt:Image size="80" resid="TicketRed80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="messageReadTaskPaneUrl" />
                  </Action>
                </Control>
                <!-- Go to http://aka.ms/ButtonCommands to learn how to add more Controls: ExecuteFunction and Menu -->
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          <!-- Go to http://aka.ms/ExtensionPointsCommands to learn how to add more Extension Points: MessageRead, AppointmentOrganizer, AppointmentAttendee -->
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="TicketRed16" DefaultValue="https://imgur.com/yfRbJKV"/>
        <bt:Image id="TicketRed32" DefaultValue="https://imgur.com/GVrV7kt"/>
        <bt:Image id="TicketRed80" DefaultValue="https://imgur.com/eH69h7A"/>
      
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="https://github.com/jcanady6/TicketSystem/blob/master/TicketSystemWeb/Functions/FunctionFile.html"/>
        <bt:Url id="messageReadTaskPaneUrl" DefaultValue="https://github.com/jcanady6/TicketSystem/blob/master/TicketSystemWeb/MessageRead.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="My Add-in Group"/>
        <bt:String id="customTabLabel"  DefaultValue="My Add-in Tab"/>
        <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties"/>
        <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties. This is an example of a button that opens a task pane."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
