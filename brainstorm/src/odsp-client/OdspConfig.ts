/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { OdspTokenConfig } from "../msal/OdspTokenConfig";
import { OdspClientProps, OdspConnectionConfig } from "./interfaces";

const connectionConfig: OdspConnectionConfig = {
	tokenProvider: new OdspTokenConfig(),
	siteUrl: "<SITE_URL>",
	driveId: "<DRIVE_ID>",
};

export const odspProps: OdspClientProps = {
	connection: connectionConfig,
};
