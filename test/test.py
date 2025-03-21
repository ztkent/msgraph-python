#!/usr/bin/env python3
import asyncio
import os
from pathlib import Path
from time import sleep
import sys
from dotenv import load_dotenv

sys.path.append(str(Path(__file__).parent.parent))

from msgraph_python.api import NewGraphAPI
from msgraph_python.exceptions import MicrosoftAuthorizationException

async def main():
    print("Loading environment variables...")
    load_dotenv()
    
    print("Starting Graph API...")
    await start(["mail", "calendar"])
    
    # Keep the event loop running
    print("Waiting for background tasks...")
    try:
        while True:
            await asyncio.sleep(1)
    except KeyboardInterrupt:
        print("\nShutting down...")

async def start(selected_scopes):
    try: 
        graph_api = await NewGraphAPI(
            client_id=os.getenv('CLIENT_ID'),
            tenant_id=os.getenv('TENANT_ID'),
            scopes=selected_scopes,
            interactive=True
        )
        print("Graph API initialized successfully")
    except MicrosoftAuthorizationException as e:
        print(f"{e}")
        return

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nProgram terminated by user")