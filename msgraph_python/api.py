import os
from azure.identity import InteractiveBrowserCredential, DeviceCodeCredential
from msgraph import GraphServiceClient
from msgraph_python.exceptions import *

# Simplify interactions with the Microsoft Graph API
# This module provides a GraphAPI class with NewGraphAPI to initialize the connection.

# The GraphAPI class provides methods to:
# - Fetch user information. More permissions will provide more information.
# - Fetch [Unread] Teams messages and Chats
# - Fetch [Unread] Outlook emails.
# - Fetch [Today's] Calendar events.

async def NewGraphAPI(client_id=None, tenant_id=None, interactive=False, scopes=["mail","calendar", "teams-chat", "teams-channel"]):
    """ Create an authenticated GraphAPI connection.
    Args:
        client_id: The client ID for the Azure app.
        tenant_id: The tenant ID for the Azure app.
        interactive: A boolean to indicate if the user should interactively authenticate.
        scopes: A list of scopes to request from the Microsoft Graph API.
    Returns:
        GraphAPI: The authenticated GraphAPI connection with the selected scopes.
    Raises:
        AuthorizationException: If the client fails to authenticate with the Microsoft Graph API.
    """
    if not client_id:
        client_id = os.getenv("CLIENT_ID")
    if not tenant_id:
        tenant_id = os.getenv("TENANT_ID")
    if not client_id or not tenant_id or not scopes:
        raise MicrosoftAuthorizationException("Invalid authentication parameters. Must provide client_id, tenant_id, and scopes.")

    selected_scopes = ['User.Read']
    if "mail" in scopes:
       selected_scopes.append('Mail.Read')
    if "calendar" in scopes:
        selected_scopes.append('Calendars.Read')
    if "teams-chat" in scopes:
        selected_scopes.append('Chat.Read')
    if "teams-channel" in scopes:
        selected_scopes.append('ChannelMessage.Read.All')
    if len(selected_scopes) == 1:
        raise MicrosoftAuthorizationException("Invalid authentication scopes. Must be 'mail' 'calendar', 'teams-chat', or 'teams-channel'.")

    if interactive:
        return GraphAPI(client=await interactive_browser_connection(selected_scopes))
    return GraphAPI(client=await device_credential_connection(client_id, tenant_id, selected_scopes))

async def device_credential_connection(client_id, tenant_id, scopes):
    # Create an application connection with the Microsoft Graph API
    credentials = DeviceCodeCredential(client_id=client_id, tenant_id=tenant_id)
    client = GraphServiceClient(credentials=credentials, scopes=scopes)
    response = await client.me.get()
    if response:
        print(response.display_name)
    else:
        raise MicrosoftAuthorizationException("Failed to authenticate user with the Microsoft Graph API")
    return client

async def interactive_browser_connection(scopes):
    # Create an application connection with the Microsoft Graph API
    # Must first be configured via preauthorization
    credentials = InteractiveBrowserCredential()
    client = GraphServiceClient(credentials=credentials, scopes=scopes)
    response = await client.me.get()
    if response:
        print(response.display_name)
    else:
        raise MicrosoftAuthorizationException("Failed to authenticate user with the Microsoft Graph API")
    return client

class GraphAPI:
    def __init__(self, client):
        """ Create a new GraphAPI object.
        Args:
            client: The GraphServiceClient object.
        """
        self.client = client
    
    # Get the user account info
    # Requires the "User.Read" permission.
    # Additional permissions will provide more information.
    async def get_user_info(self):
        """ 
        Get the user account info.
        Permissions:
            User.Read
        Returns:
            A dictionary with the following keys:
                additional_data, id, odata_type, deleted_date_time, about_me, account_enabled, activities, age_group, 
                agreement_acceptances, app_role_assignments, assigned_licenses, assigned_plans, authentication, 
                authorization_info, birthday, business_phones, calendar, calendar_groups, calendar_view, calendars, 
                chats, city, company_name, consent_provided_for_minor, contact_folders, contacts, country, 
                created_date_time, created_objects, creation_type, custom_security_attributes, department, 
                device_enrollment_limit, device_management_troubleshooting_events, direct_reports, display_name, 
                drive, drives, employee_experience, employee_hire_date, employee_id, employee_leave_date_time, 
                employee_org_data, employee_type, events, extensions, external_user_state, 
                external_user_state_change_date_time, fax_number, followed_sites, given_name, hire_date, identities, 
                im_addresses, inference_classification, insights, interests, is_resource_account, job_title, 
                joined_teams, last_password_change_date_time, legal_age_group_classification, 
                license_assignment_states, license_details, mail, mail_folders, mail_nickname, mailbox_settings, 
                managed_app_registrations, managed_devices, manager, member_of, messages, mobile_phone, my_site, 
                oauth2_permission_grants, office_location, on_premises_distinguished_name, on_premises_domain_name, 
                on_premises_extension_attributes, on_premises_immutable_id, on_premises_last_sync_date_time, 
                on_premises_provisioning_errors, on_premises_sam_account_name, on_premises_security_identifier, 
                on_premises_sync_enabled, on_premises_user_principal_name, onenote, online_meetings, other_mails, 
                outlook, owned_devices, owned_objects, password_policies, password_profile, past_projects, people, 
                permission_grants, photo, photos, planner, postal_code, preferred_data_location, preferred_language, 
                preferred_name, presence, print, provisioned_plans, proxy_addresses, registered_devices, 
                responsibilities, schools, scoped_role_member_of, security_identifier, service_provisioning_errors, 
                settings, show_in_address_list, sign_in_activity, sign_in_sessions_valid_from_date_time, skills, 
                state, street_address, surname, teamwork, todo, transitive_member_of, usage_location, 
                user_principal_name, user_type
        Raises:
            RequestException: If the request to get user info fails.
        """
        try: 
            response = await self.client.me.get()
            response_dict = response.__dict__
            return response_dict
        except Exception as e:
            raise MicrosoftRequestException(f"Failed to get user info: {e}")

    # Get messages from the teams channels the authorized user is a part of
    # Requires the "ChannelMessage.Read.All" permission. 
    # This requires admin consent.
    async def get_teams_channel_messages(self):
        """ 
        Get messages from the teams channels the authorized user is a part of.
        Permissions:
            ChannelMessage.Read.All
        Returns:
            A dictionary where the keys are team IDs and the values are lists of messages in the channels of those teams.
        Raises:
            RequestException: If the request to get Teams messages fails.
        """
        try: 
            teams = await self.client.me.joined_teams.get()
            messages = {}
            for team in teams:
                team_messages = []
                channels = await self.client.teams[team.id].channels.get()
                for channel in channels:
                    channel_messages = await self.client.teams[team.id].channels[channel.id].messages.get()
                    team_messages.extend(channel_messages)
                messages[team.id] = team_messages
            return messages
        except Exception as e:
            raise MicrosoftRequestException(f"Failed to get Teams messages: {e}")

    # Get the unread messages from the teams the authorized user is a part of
    # Requires the "ChannelMessage.Read.All" permission
    # This requires admin consent.
    async def get_unread_teams_channel_messages(self):
        """ 
        Get unread messages from the teams channels the authorized user is a part of.
        Permissions:
            ChannelMessage.Read.All
        Returns:
            A dictionary where the keys are team IDs and the values are lists of unread messages in the channels of those teams.
        Raises:
            RequestException: If the request to get Teams messages fails.
        """
        try: 
            teams = await self.client.me.joined_teams.get()
            messages = {}
            for team in teams:
                team_messages = []
                channels = await self.client.teams[team.id].channels.get()
                for channel in channels:
                    channel_messages = await self.client.teams[team.id].channels[channel.id].messages.request().filter("isRead eq false").get()
                    team_messages.extend(channel_messages)
                messages[team.id] = team_messages
            return messages
        except Exception as e:
            raise MicrosoftRequestException(f"Failed to get Teams messages: {e}")

    # Get the messages from a specific chat
    # Requires the "Chat.Read" permission
    async def get_teams_chat_messages(self, chat_id):
        """ 
        Get the Teams messages from a specific chat.
        Permissions:
            Chat.Read
        Returns:
            A list of messages from the specified chat.
        Raises:
            RequestException: If the request to get chat messages fails.
        """
        try: 
            response = await self.client.me.chats[chat_id].messages.get()
            return response
        except Exception as e:
            raise MicrosoftRequestException(f"Failed to get chat messages: {e}")

    # Get all the chats of the authorized user
    # Requires the "Chat.Read" permission
    async def get_all_teams_chats(self):
        """ 
        Get all the Teams chats and messages of the authorized user.
        Permissions:
            Chat.Read
        Returns:
            A list of all chats of the authorized user.
        Raises:
            RequestException: If the request to get chats fails.
        """
        try: 
            response = await self.client.me.chats.get()
            return response
        except Exception as e:
            raise MicrosoftRequestException(f"Failed to get chats: {e}")
    
    # Get all unread messages from all chats of the authorized user
    # Requires the "Chat.Read" permission
    async def get_all_unread_teams_chat_messages(self):
        """ 
        Get all unread Teams messages from all Teams chats of the authorized user.
        Permissions:
            Chat.Read
        Returns:
            A dictionary where the keys are chat IDs and the values are lists of unread messages in those chats.
        Raises:
            RequestException: If the request to get chat messages fails.
        """
        try: 
            chats = await self.client.me.chats.get()
            messages = {}
            for chat in chats:
                chat_messages = await self.client.me.chats[chat.id].messages.get()
                unread_messages = [message for message in chat_messages if message.isRead == False]
                messages[chat.id] = unread_messages
            return messages
        except Exception as e:
            raise MicrosoftRequestException(f"Failed to get chat messages: {e}")

    # Get the emails of the authorized user
    # Requires the "Mail.Read" permission
    async def get_outlook_emails(self):
        """ 
        Get the emails of the authorized user.
        Permissions:
            Mail.Read
        Returns:
            A list of emails of the authorized user.
        Raises:
            RequestException: If the request to get emails fails.
        """
        try: 
            response = await self.client.me.messages.get()
            return response
        except Exception as e:
            raise MicrosoftRequestException(f"Failed to get emails: {e}")

    # Get the unread emails of the authorized user from the Outlook inbox
    # Requires the "Mail.Read" permission
    async def get_unread_outlook_emails(self):
        """ 
        Get the unread emails of the authorized user from the Outlook inbox.
        Permissions:
            Mail.Read
        Returns:
            A list of unread emails of the authorized user.
        Raises:
            RequestException: If the request to get emails fails.
        """
        try: 
            response = await self.client.me.mail_folders.get()
            return response
        except Exception as e:
            raise MicrosoftRequestException(f"Failed to get emails: {e}")
        
    # Get all the calendar events of the authorized user
    # Requires the "Calendars.Read" permission
    async def get_all_calendar_events(self):
        """ 
        Get all the calendar events of the authorized user.
        Permissions:
            Calendars.Read
        Returns:
            A list of all calendar events of the authorized user.
        Raises:
            RequestException: If the request to get calendar events fails.
        """
        try: 
            response = await self.client.me.events.get()
            return response
        except Exception as e:
            raise MicrosoftRequestException(f"Failed to get calendar events: {e}")

    # Get the calendar events of the authorized user for today
    # Requires the "Calendars.Read" permission
    async def get_todays_calendar_events(self):
        """ 
        Get the calendar events of the authorized user for today.
        Permissions:
            Calendars.Read
        Returns:
            A list of calendar events of the authorized user for today.
        Raises:
            RequestException: If the request to get calendar events fails.
        """
        try: 
            response = await self.client.me.calendar_view.get()
            return response
        except Exception as e:
            raise MicrosoftRequestException(f"Failed to get todays calendar events: {e}")