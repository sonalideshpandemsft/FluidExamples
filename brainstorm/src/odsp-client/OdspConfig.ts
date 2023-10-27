/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { getTokens } from "../msal/tokens";
import { OdspClientProps, OdspConnectionConfig, tokenMap } from "./interfaces";
import { OdspClient } from "./OdspClient";

export const odspConfig = async () => {
	console.log("Authenticating------");

	await getTokens();

	const getStorageToken = async () => {
		return tokenMap.get("sharePointToken");
	};

	const getPushToken = async () => {
		return tokenMap.get("pushToken");
	};

	const connectionConfig: OdspConnectionConfig = {
		getSharePointToken: getStorageToken as any,
		getPushServiceToken: getPushToken as any,
		siteUrl: "<SITE_URL>",
		driveId: "<DRIVE_ID>",
	};

	const conn: OdspClientProps = {
		connection: connectionConfig,
	};

	OdspClient.init(conn);
	return conn;
};
