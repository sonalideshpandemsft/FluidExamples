/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */
import { ServiceAudience } from "@fluidframework/fluid-static";
import { type IClient } from "@fluidframework/protocol-definitions";

import { type OdspMember, type IOdspAudience, OdspUser } from "./interfaces";

export class OdspAudience extends ServiceAudience<OdspMember> implements IOdspAudience {
	protected createServiceMember(audienceMember: IClient): OdspMember {
		const user = audienceMember.user as OdspUser;
		if (user?.name === undefined) {
			throw new Error("Provided user was not an OdspUser");
		}

		return {
			userId: user.id,
			userName: user.name,
			email: user.email,
			connections: [],
		};
	}
}
