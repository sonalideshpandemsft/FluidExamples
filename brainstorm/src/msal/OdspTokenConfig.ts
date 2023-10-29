import { type ITokenProvider, type ITokenResponse } from "@fluidframework/routerlicious-driver";
import { tokenMap } from "../odsp-client";

export class OdspTokenConfig implements ITokenProvider {
	// public constructor(siteUrl: string, itemId?: string) {}

	public async fetchOrdererToken(tenantId: string, documentId?: string): Promise<ITokenResponse> {
		return {
			jwt: tokenMap.get("pushToken") as string,
		};
	}

	public async fetchStorageToken(tenantId: string, documentId: string): Promise<ITokenResponse> {
		return {
			jwt: tokenMap.get("sharePointToken") as string,
		};
	}
}
