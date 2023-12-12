import { TokenResponse } from "@fluidframework/odsp-driver-definitions";
import { IOdspTokenProvider } from "@fluid-experimental/odsp-client";
import { tokenMap } from "../odsp-client/OdspConfig";

export class OdspTokenConfig implements IOdspTokenProvider {
	// public constructor(siteUrl: string, itemId?: string) {}

	public async fetchWebsocketToken(siteUrl: string, refresh: boolean): Promise<TokenResponse> {
		return {
			token: tokenMap.get("pushToken") as string,
		};
	}

	public async fetchStorageToken(siteUrl: string, refresh: boolean): Promise<TokenResponse> {
		return {
			token: tokenMap.get("sharePointToken") as string,
		};
	}
}
