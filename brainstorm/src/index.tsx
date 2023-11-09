/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */
import { initializeIcons, ThemeProvider } from "@fluentui/react";
import { IFluidContainer } from "@fluidframework/fluid-static";
import React from "react";
import ReactDOM from "react-dom";
import { BrainstormView } from "./view/BrainstormView";
import "./view/index.css";
import "./view/App.css";
import { themeNameToTheme } from "./view/Themes";
import { containerSchema } from "./Config";
import { OdspContainerServices } from "./odsp-client/interfaces";
import { OdspClient } from "./odsp-client/OdspClient";
import { ConnectionState } from "fluid-framework";
import { getTokens } from "./msal/tokens";
import { odspProps } from "./odsp-client";

export async function start() {
	initializeIcons();

	await getTokens();

	console.log("-----TOKENS GENERATED----");

	const client = new OdspClient(odspProps);

	const getContainerId = (): { containerId: string; isNew: boolean } => {
		let isNew = false;
		if (location.hash.length === 0) {
			isNew = true;
		}
		const hash = location.hash;
		const absoluteUrl = hash.charAt(0) === "#" ? hash.substring(1) : hash;
		const containerId = localStorage.getItem(absoluteUrl) as string;
		return { containerId, isNew };
	};

	const { containerId, isNew } = getContainerId();

	let container: IFluidContainer;
	let services: OdspContainerServices;

	if (isNew) {
		console.log("CREATING THE CONTAINER");

		({ container, services } = await client.createContainer(containerSchema));

		const itemId = await container.attach();
		const attributes = await services.tenantAttributes();
		if (attributes === undefined) {
			throw new Error("Tenant attributes are undefined");
		}
		const sharingUrl = attributes.sharingUrl;
		if (itemId === undefined || sharingUrl === undefined) {
			throw new Error("Either itemId or url is undefined");
		}

		localStorage.setItem(itemId, sharingUrl);
		location.hash = itemId;
	} else {
		console.log("GET CONTAINER", containerId);
		({ container, services } = await client.getContainer(containerId, containerSchema));
	}

	if (container.connectionState !== ConnectionState.Connected) {
		await new Promise<void>((resolve) => {
			container.once("connected", () => {
				resolve();
			});
		});
	}

	ReactDOM.render(
		<React.StrictMode>
			<ThemeProvider theme={themeNameToTheme("default")}>
				<BrainstormView container={container} services={services} />
			</ThemeProvider>
		</React.StrictMode>,
		document.getElementById("root"),
	);
}

start().catch((error) => console.error(error));
