class AuthorizationException(Exception):
    """Exception raised when failing to connect to the Microsoft Graph API."""

    def __init__(self, message="Failed to connect to the Microsoft Graph API"):
        self.message = message
        super().__init__(self.message)


class RequestException(Exception):
    """Exception raised when a Microsoft Graph API request fails."""

    def __init__(self, message="Microsoft Graph API request failed"):
        self.message = message
        super().__init__(self.message)