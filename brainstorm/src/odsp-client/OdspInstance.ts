/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { Container, Loader } from "@fluidframework/container-loader";
import {
	AttachState,
	IContainer,
	IFluidModuleWithDetails,
	IRuntimeFactory,
} from "@fluidframework/container-definitions";
import { IDocumentServiceFactory } from "@fluidframework/driver-definitions";
import {
	OdspDocumentServiceFactory,
	createOdspCreateContainerRequest,
	OdspDriverUrlResolver,
} from "@fluidframework/odsp-driver";
import {
	ContainerSchema,
	DOProviderContainerRuntimeFactory,
	FluidContainer,
} from "@fluidframework/fluid-static";
import {
	OdspCreateContainerConfig,
	OdspGetContainerConfig,
	OdspResources,
	OdspClientProps,
} from "./interfaces";
import { OdspAudience } from "./OdspAudience";
import { IClient } from "@fluidframework/protocol-definitions";
import { IOdspResolvedUrl } from "@fluidframework/odsp-driver-definitions";

/**
 * OdspInstance provides the ability to have a Fluid object backed by the ODSP service
 */
export class OdspInstance {
	public readonly documentServiceFactory: IDocumentServiceFactory;
	public readonly urlResolver: OdspDriverUrlResolver;

	constructor(private readonly properties: OdspClientProps) {
		this.documentServiceFactory = new OdspDocumentServiceFactory(
			this.properties.connection.getSharePointToken,
			this.properties.connection.getPushServiceToken,
		);

		this.urlResolver = new OdspDriverUrlResolver();
	}

	public async createContainer(
		serviceContainerConfig: OdspCreateContainerConfig,
		containerSchema: ContainerSchema,
	): Promise<OdspResources> {
		const container = await this.getContainerInternal(
			serviceContainerConfig,
			new DOProviderContainerRuntimeFactory(containerSchema),
			true,
		);

		return this.getContainerAndServices(container);
	}

	public async getContainer(
		serviceContainerConfig: OdspGetContainerConfig,
		containerSchema: ContainerSchema,
	): Promise<OdspResources> {
		console.log("load container");
		const container = await this.getContainerInternal(
			serviceContainerConfig,
			new DOProviderContainerRuntimeFactory(containerSchema),
			false,
		);

		return this.getContainerAndServices(container);
	}

	public async containerPath(url: string) {
		const itemIdPattern = /itemId=([^&]+)/; // regular expression to match the itemId parameter value
		let itemId;

		const match = url.match(itemIdPattern); // get the match object for the itemId parameter value
		if (match) {
			itemId = match[1]; // extract the itemId parameter value from the match object
			console.log(itemId); // output: "itemidQ"
		} else {
			console.log("itemId parameter not found in the URL");
			itemId = "";
		}
		return itemId;
	}

	private async getContainerAndServices(container: IContainer): Promise<OdspResources> {
		const rootDataObject = (await container.request({ url: "/" })).value;
		const fluidContainer = new FluidContainer(container, rootDataObject);
		const resolvedUrl = container.resolvedUrl as IOdspResolvedUrl;

		const containerServices = {
			getSharingUrl: async () => {
				const url = await container.getAbsoluteUrl("/");
				if (url === undefined) {
					throw new Error("container has no url");
				}
				return url;
			},
			getItemId: async () => {
				return resolvedUrl.itemId;
			},
			getContainerId: async () => {
				return resolvedUrl.id;
			},
			audience: new OdspAudience(container),
		};

		const odspContainerServices: OdspResources = { fluidContainer, containerServices };

		return odspContainerServices;
	}

	private async getContainerInternal(
		containerConfig: OdspCreateContainerConfig | OdspGetContainerConfig,
		containerRuntimeFactory: IRuntimeFactory,
		createNew: boolean,
	): Promise<IContainer> {
		const load = async (): Promise<IFluidModuleWithDetails> => {
			return {
				module: { fluidExport: containerRuntimeFactory },
				details: { package: "no-dynamic-package", config: {} },
			};
		};

		const codeLoader = { load };

		const client: IClient = {
			details: {
				capabilities: { interactive: true },
			},
			permission: [],
			scopes: [],
			user: { id: "" },
			mode: "write",
		};

		console.log("resolver: ", this.urlResolver);
		const loader = new Loader({
			urlResolver: this.urlResolver,
			documentServiceFactory: this.documentServiceFactory,
			codeLoader,
			logger: containerConfig.logger,
			options: { client },
		});

		let container: IContainer;
		if (createNew) {
			// Generate an ODSP driver specific new file request using the provided metadata for the file from the
			// containerConfig.
			const { siteUrl, driveId, folderName, fileName } =
				containerConfig as OdspCreateContainerConfig;

			const request = createOdspCreateContainerRequest(
				siteUrl,
				driveId,
				folderName,
				fileName,
			);
			// We're not actually using the code proposal (our code loader always loads the same module regardless of the
			// proposal), but the Container will only give us a NullRuntime if there's no proposal.  So we'll use a fake
			// proposal.
			container = (await loader.createDetachedContainer({
				package: "",
				config: {},
			})) as Container;
			if (container.attachState !== AttachState.Detached) {
				throw new Error("Cannot attach container. Container is not in detached state");
			}
			await container.attach(request);
			if (container.resolvedUrl === undefined) {
				throw new Error("Cannot attach container. Container is not in detached state");
			}
		} else {
			// Generate the request to fetch our existing container back using the provided SharePoint
			// file url. If this is a share URL, it needs to be redeemed by the service to be accessible
			// by other users. As such, we need to set the appropriate header for those scenarios.
			const { fileUrl } = containerConfig as OdspGetContainerConfig;
			const request = {
				url: fileUrl,
			};
			// Request must be appropriate and parseable by resolver.
			container = await loader.resolve(request);
		}
		return container;
	}
}
