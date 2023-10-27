/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { v4 as uuid } from "uuid";
import { initializeIcons, ThemeProvider } from "@fluentui/react";
import { ConnectionState, FluidContainer } from "fluid-framework";
import React from "react";
import ReactDOM from "react-dom";
import { BrainstormView } from "./view/BrainstormView";
import "./view/index.css";
import "./view/App.css";
import { themeNameToTheme } from "./view/Themes";
import { containerSchema } from "./Config";
import {
	OdspContainerServices,
	OdspCreateContainerConfig,
	OdspGetContainerConfig,
} from "./odsp-client/interfaces";
import { OdspClient } from "./odsp-client/OdspClient";
import { odspConfig } from "./odsp-client";

export async function start() {
	initializeIcons();

	const odspDriver = await odspConfig();

	const getContainerId = (): { containerId: string; isNew: boolean } => {
		let isNew = false;
		console.log("hash: ", location.hash);
		if (location.hash.length === 0) {
			isNew = true;
		}
		const hash = location.hash;
		const itemId = hash.charAt(0) === "#" ? hash.substring(1) : hash;
		const containerId = localStorage.getItem(itemId) as string;
		return { containerId, isNew };
	};

	const { containerId, isNew } = getContainerId();

	let container: FluidContainer;
	let services: OdspContainerServices;

	if (isNew) {
		console.log("CREATING THE CONTAINER");
		const containerConfig: OdspCreateContainerConfig = {
			siteUrl: odspDriver.connection.siteUrl,
			driveId: odspDriver.connection.driveId,
			folderName: "",
			fileName: uuid(),
		};

		console.log("CONTAINER CONFIG", containerConfig);

		const { fluidContainer, containerServices } = await OdspClient.createContainer(
			containerConfig,
			containerSchema,
		);
		container = fluidContainer;
		services = containerServices;

		const url = await containerServices.getSharingUrl();
		const containerId = await containerServices.getContainerId();
		localStorage.setItem(containerId, url);
		console.log("CONTAINER CREATED");
		location.hash = containerId;
	} else {
		const containerConfig: OdspGetContainerConfig = {
			fileUrl: containerId, //pass file url
		};

		const { fluidContainer, containerServices } = await OdspClient.getContainer(
			containerConfig,
			containerSchema,
		);

		container = fluidContainer;
		services = containerServices;
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
