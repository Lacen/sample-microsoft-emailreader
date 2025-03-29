import sys
import json
import requests
import time


def main():
    """
    Main Function for execution
    """

    with open("config.json", "r") as f:
        config = json.load(f)
    if "access_token" not in config:
        access_token = authorize_user(config)
    else:
        access_token = refresh_access_token(config)
    read_outlook_emails(config["user_principal_name"], access_token)


def get_token(config: dict, device_code: str) -> str:
    """Function for getting Access Token and Refresh Token
    Args:
            config (dict): The JSON file containing the users configuration
            device_code (str): The Authorization code from User Authentication

    Returns:
            token (str): The access_token
    """

    client_id = config["client_id"]
    tenant_id = config["tenant_id"]
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    token_payload = {
        "client_id": client_id,
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "device_code": device_code,
    }

    # Waits for User to Validate through the Portal
    while True:
        token_response = requests.post(token_url, data=token_payload)
        token_data = token_response.json()

        if "access_token" in token_data:
            print("Access Token Received!")
            ACCESS_TOKEN = token_data["access_token"]
            REFRESH_TOKEN = token_data["refresh_token"]

            # Store token information to be used next time without authenticating again
            with open("config.json", "w") as f:
                config["access_token"] = ACCESS_TOKEN
                config["refresh_token"] = REFRESH_TOKEN
                json.dump(config, f, indent=4)

            return ACCESS_TOKEN

        elif "error" in token_data and token_data["error"] == "authorization_pending":
            print("Waiting for user to sign in...")
            time.sleep(5)

        else:
            sys.exit("Error:", token_data)


def authorize_user(config: dict) -> str:
    """Authorize a User

    Args:
        config (dict): The JSON file containing the users configuration

    Returns:
        access_token = the access token

    """

    client_id = config["client_id"]
    tenant_id = config["tenant_id"]
    scope = config["scope"]
    auth_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode"

    data = {"client_id": client_id, "scope": scope}
    response = requests.post(auth_url, data=data)
    data = response.json()

    device_code = data["device_code"]
    print(
        f" Login in at {data['verification_uri']} and enter the code: {data['user_code']}"
    )
    access_token = get_token(config=config, device_code=device_code)
    return access_token


def refresh_access_token(config: dict):
    """Refreshes the access token to be used

    Args:
        config (dict): The JSON file containing the users configuration

    Returns:
        access_token: The access token
    """

    client_id = config["client_id"]
    tenant_id = config["tenant_id"]
    scope = config["scope"]
    token = config["refresh_token"]
    refresh_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "refresh_token",
        "client_id": client_id,
        "refresh_token": token,
        "scope": scope,
    }
    response = requests.post(refresh_url, data=data)

    tokens = response.json()
    new_access_token = tokens.get("access_token")
    new_refresh_token = tokens.get("refresh_token")
    with open("config.json", "w") as f:
        config["access_token"] = new_access_token
        config["refresh_token"] = new_refresh_token
        json.dump(config, f, indent=4)
    return new_access_token


def read_outlook_emails(user_principal_name: str, token: str):
    url = f"https://graph.microsoft.com/v1.0/users/{user_principal_name}/messages"
    headers = {
        "Authorization": f"Bearer {token}",
    }
    response = requests.get(url, headers=headers)
    for message in response.json()["value"]:
        sender_email = message["sender"]["emailAddress"]["address"]
        body_preview = message["bodyPreview"]

        print(f"Sender: {sender_email}")
        print(f"{body_preview.strip()}\n\n\n")


main()
