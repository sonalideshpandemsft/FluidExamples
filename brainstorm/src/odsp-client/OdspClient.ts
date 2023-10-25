/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { Loader } from "@fluidframework/container-loader";
import {
	AttachState,
	IContainer,
	IFluidModuleWithDetails,
} from "@fluidframework/container-definitions";
import { IDocumentServiceFactory, IUrlResolver } from "@fluidframework/driver-definitions";
import {
	OdspDocumentServiceFactory,
	createOdspCreateContainerRequest,
} from "@fluidframework/odsp-driver";
import {
	ContainerSchema,
	DOProviderContainerRuntimeFactory,
	FluidContainer,
	IFluidContainer,
} from "@fluidframework/fluid-static";
import {
	OdspClientProps,
	OdspContainerServices,
	OdspConnectionConfig,
} from "./interfaces";
import { OdspAudience } from "./OdspAudience";
import { OdspUrlResolver } from "./odspUrlResolver";
import { IClient } from "@fluidframework/protocol-definitions";
import { OdspResourceTokenFetchOptions, TokenFetcher } from "@fluidframework/odsp-driver-definitions";

/**
 * OdspClient provides the ability to have a Fluid object backed by the ODSP service
 */
export class OdspClient {
	private readonly documentServiceFactory: IDocumentServiceFactory;
	private readonly urlResolver: IUrlResolver;
	private readonly configProvider: IConfigProviderBase | undefined;

	constructor(private readonly properties: OdspClientProps) {
		this.urlResolver = new OdspUrlResolver();

		const getSharePointToken = async (options: OdspResourceTokenFetchOptions) => {
			const tokenResponse = await this.properties.connection.tokenProvider.fetchStorageToken(options.tenantId as any, options.itemId as any);
			// Extract the token property from the response
			const token = tokenResponse.jwt;
			return token;
		};
		
		const getWebsocketToken = async (options: OdspResourceTokenFetchOptions) => {
			const tokenResponse = await this.properties.connection.tokenProvider.fetchOrdererToken(options.tenantId as any, options.itemId as any);
			// Extract the token property from the response
			const token = tokenResponse.jwt;
			return token;
		};
		
		const originalDocumentServiceFactory = new OdspDocumentServiceFactory(
			getSharePointToken,
			getWebsocketToken,
		);
		
		this.documentServiceFactory = applyStorageCompression(
			originalDocumentServiceFactory,
			properties.summaryCompression,
		);
		this.configProvider = properties.configProvider;
	}

	async getContainer(
		url: string,
		containerSchema: ContainerSchema,
	): Promise<{
		container: IFluidContainer;
		services: OdspContainerServices;}> {
		const loader = this.createLoader(containerSchema);

		// Request must be appropriate and parseable by resolver.
		const container = await loader.resolve({ url });
		const rootDataObject = await requestFluidObject<IRootDataObject>(container, "/");
		const fluidContainer = new FluidContainer(container, rootDataObject);

		const services = this.getContainerServices(container);
		return { container: fluidContainer, services }
	}

	private getContainerServices(container: IContainer): OdspContainerServices {
		return {
			audience: new OdspAudience(container),
		};
	}

	private createLoader(containerSchema: ContainerSchema): Loader {
		const runtimeFactory = new DOProviderContainerRuntimeFactory(containerSchema);
		const load = async (): Promise<IFluidModuleWithDetails> => {
			return {
				module: { fluidExport: runtimeFactory },
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

		return new Loader({
			urlResolver: this.urlResolver,
			documentServiceFactory: this.documentServiceFactory,
			codeLoader,
			logger: this.properties.logger,
			options: { client },
			configProvider: this.configProvider,
		});
	}

	public async createContainer(containerSchema: ContainerSchema): Promise<{
		container: IFluidContainer;
		services: OdspContainerServices;
	}> {
		const loader = this.createLoader(containerSchema);

		const container = await loader.createDetachedContainer({
			package: "no-dynamic-package",
			config: {}
		});

		const fluidContainer = await this.createFluidContainer(container, this.properties.connection);

		const services = this.getContainerServices(container);

		return { container: fluidContainer, services };
	}

	private async createFluidContainer(
		container: IContainer,
		connection: OdspConnectionConfig,
	): Promise<FluidContainer> {
		const createNewRequest = createOdspCreateContainerRequest(
			connection.endpoint,
			connection.driveId,
			connection.folderName,
			connection.fileName,
		);

		// eslint-disable-next-line import/no-deprecated
		const rootDataObject = await requestFluidObject<IRootDataObject>(container, "/");

		/**
		 * See {@link FluidContainer.attach}
		 */
		const attach = async (): Promise<string> => {
			if (container.attachState !== AttachState.Detached) {
				throw new Error("Cannot attach container. Container is not in detached state");
			}
			await container.attach(createNewRequest);
			if (container.resolvedUrl === undefined) {
				throw new Error("Resolved Url not available on attached container");
			}
			return container.resolvedUrl.id;
		};
		const fluidContainer = new FluidContainer(container, rootDataObject);
		fluidContainer.attach = attach;
		return fluidContainer;
	}
}
