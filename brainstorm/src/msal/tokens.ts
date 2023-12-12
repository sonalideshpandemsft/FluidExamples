/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import {
	PublicClientApplication,
	AuthenticationResult,
	InteractionRequiredAuthError,
} from "@azure/msal-browser";
import { tokenMap } from "../odsp-client/OdspConfig";

const msalConfig = {
	auth: {
		clientId: "<APP-ID>",
		authority: "https://login.microsoftonline.com/common/",
	},
};

const graphScopes = ["FileStorageContainer.Selected"];

const sharePointScopes = ["https://<TENANT-NAME>.sharepoint.com/Container.Selected"];

const pushScopes = ["offline_access", "https://pushchannel.1drv.ms/PushChannel.ReadWrite.All"];

const msalInstance = new PublicClientApplication(msalConfig);

export async function getTokens(): Promise<{
	graphToken: string;
}> {
	const response = await msalInstance.loginPopup({ scopes: graphScopes });

	msalInstance.setActiveAccount(response.account);

	try {
		// Attempt to acquire SharePoint token silently
		const sharePointRequest = {
			scopes: sharePointScopes,
		};
		const sharePointTokenResult: AuthenticationResult = await msalInstance.acquireTokenSilent(
			sharePointRequest,
		);

		// Attempt to acquire other token silently
		const otherRequest = {
			scopes: pushScopes,
		};
		const pushTokenResult: AuthenticationResult = await msalInstance.acquireTokenSilent(
			otherRequest,
		);

		tokenMap.set("sharePointToken", sharePointTokenResult.accessToken);
		tokenMap.set("pushToken", pushTokenResult.accessToken);

		// Return both tokens
		return {
			graphToken: response.accessToken,
		};
	} catch (error) {
		if (error instanceof InteractionRequiredAuthError) {
			// If silent token acquisition fails, fall back to interactive flow
			const sharePointRequest = {
				scopes: sharePointScopes,
			};
			const sharePointTokenResult: AuthenticationResult =
				await msalInstance.acquireTokenPopup(sharePointRequest);

			const otherRequest = {
				scopes: pushScopes,
			};
			const pushTokenResult: AuthenticationResult = await msalInstance.acquireTokenPopup(
				otherRequest,
			);

			tokenMap.set("sharePointToken", sharePointTokenResult.accessToken);
			tokenMap.set("pushToken", pushTokenResult.accessToken);

			// Return both tokens
			return {
				graphToken: response.accessToken,
			};
		} else {
			// Handle any other error
			console.error(error);
			throw error;
		}
	}
}
