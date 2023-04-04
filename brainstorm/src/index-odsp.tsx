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
import { getodspDriver } from "./odsp-client";

const documentId = uuid();

export async function start() {
	initializeIcons();

	console.log("Initiating the driver------");
	const odspDriver = await getodspDriver();
	console.log("INITIAL DRIVER", odspDriver);

	const getContainerId = (): { containerId: string; isNew: boolean } => {
		let isNew = false;
		if (location.hash) {
			isNew = true;
		}
		const hash = decodeURI(location.hash);
		const containerId = hash.charAt(0) === "#" ? hash.substring(1) : hash;
		return { containerId, isNew };
	};

	const { containerId, isNew } = getContainerId();

	let container: FluidContainer;
	let services: OdspContainerServices;

	if (isNew) {
		console.log("CREATING THE CONTAINER");
		const containerConfig: OdspCreateContainerConfig = {
			siteUrl: odspDriver.siteUrl,
			driveId: odspDriver.driveId,
			folderName: odspDriver.directory,
			fileName: documentId,
		};

		console.log("CONTAINER CONFIG", containerConfig);

		const { fluidContainer, containerServices } = await OdspClient.createContainer(
			containerConfig,
			containerSchema,
		);
		container = fluidContainer;
		services = containerServices;

		const sharingLink = await containerServices.generateLink();

		console.log("CONTAINER CREATED");
		location.hash = encodeURI(sharingLink);
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
