/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
import { OdspClientProps, OdspConnectionConfig } from "@fluid-experimental/odsp-client";
import { OdspTokenConfig } from "../msal/OdspTokenConfig";

const connectionConfig: OdspConnectionConfig = {
	tokenProvider: new OdspTokenConfig(),
	siteUrl: "<SITE_URL>",
	driveId: "<RAAS_DRIVE_ID>",
};

export const odspProps: OdspClientProps = {
	connection: connectionConfig,
};

export const tokenMap: Map<string, string> = new Map();
