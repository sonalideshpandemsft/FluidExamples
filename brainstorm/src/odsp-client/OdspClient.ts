import { IDocumentServiceFactory } from "@fluidframework/driver-definitions";
import {
	OdspDocumentServiceFactory,
	OdspDriverUrlResolver,
	createOdspCreateContainerRequest,
} from "@fluidframework/odsp-driver";
import {
	OdspClientProps,
	OdspConnectionConfig,
	OdspContainerServices,
	OdspServiceAttributes,
} from "./interfaces";
import {
	type ContainerSchema,
	DOProviderContainerRuntimeFactory,
	IFluidContainer,
	FluidContainer,
} from "@fluidframework/fluid-static";
import {
	AttachState,
	IContainer,
	IFluidModuleWithDetails,
} from "@fluidframework/container-definitions";
import { IClient } from "@fluidframework/protocol-definitions";
import { Loader } from "@fluidframework/container-loader";
import { v4 as uuid } from "uuid";
import {
	IOdspResolvedUrl,
	OdspResourceTokenFetchOptions,
} from "@fluidframework/odsp-driver-definitions";
import { OdspAudience } from "./OdspAudience";
import { ITokenResponse } from "@fluidframework/azure-client";

export class OdspClient {
	private documentServiceFactory: IDocumentServiceFactory;
	private urlResolver: OdspDriverUrlResolver;

	public constructor(private readonly properties: OdspClientProps) {
		const getSharePointToken = async (options: OdspResourceTokenFetchOptions) => {
			const tokenResponse: ITokenResponse =
				await this.properties.connection.tokenProvider.fetchStorageToken(
					options.siteUrl,
					"",
				);
			return {
				token: tokenResponse.jwt,
			};
		};

		const getPushServiceToken = async (options: OdspResourceTokenFetchOptions) => {
			const tokenResponse: ITokenResponse =
				await this.properties.connection.tokenProvider.fetchOrdererToken(options.siteUrl);
			return {
				token: tokenResponse.jwt,
			};
		};
		this.documentServiceFactory = new OdspDocumentServiceFactory(
			getSharePointToken,
			getPushServiceToken,
		);

		this.urlResolver = new OdspDriverUrlResolver();
	}

	public async createContainer(containerSchema: ContainerSchema): Promise<{
		container: IFluidContainer;
		services: OdspContainerServices;
	}> {
		const loader = this.createLoader(containerSchema);

		const container = await loader.createDetachedContainer({
			package: "no-dynamic-package",
			config: {},
		});

		const fluidContainer = await this.createFluidContainer(
			container,
			this.properties.connection,
		);

		const services = await this.getContainerServices(container);

		return { container: fluidContainer, services };
	}

	public async getContainer(
		sharingUrl: string,
		containerSchema: ContainerSchema,
	): Promise<{
		container: IFluidContainer;
		services: OdspContainerServices;
	}> {
		const loader = this.createLoader(containerSchema);
		const container = await loader.resolve({ url: sharingUrl });

		const rootDataObject = (await container.request({ url: "/" })).value;
		const fluidContainer = new FluidContainer(container, rootDataObject);
		const services = await this.getContainerServices(container);
		return { container: fluidContainer, services };
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
		});
	}

	private async createFluidContainer(
		container: IContainer,
		connection: OdspConnectionConfig,
	): Promise<IFluidContainer> {
		const createNewRequest = createOdspCreateContainerRequest(
			connection.siteUrl,
			connection.driveId,
			connection.path,
			uuid(),
		);

		const rootDataObject = (await container.request({ url: "/" })).value;

		/**
		 * See {@link FluidContainer.attach}
		 */
		const attach = async (): Promise<string> => {
			if (container.attachState !== AttachState.Detached) {
				throw new Error("Cannot attach container. Container is not in detached state");
			}
			await container.attach(createNewRequest);
			const resolvedUrl = container.resolvedUrl as IOdspResolvedUrl;
			console.log("create: ", resolvedUrl.itemId);
			if (container.resolvedUrl === undefined) {
				throw new Error("Resolved Url not available on attached container");
			}
			return resolvedUrl.url;
		};
		const fluidContainer = new FluidContainer(container, rootDataObject);
		fluidContainer.attach = attach;
		return fluidContainer;
	}

	private async getContainerServices(container: IContainer): Promise<OdspContainerServices> {
		const getAttributes = async (): Promise<OdspServiceAttributes> => {
			const resolvedUrl = container.resolvedUrl as IOdspResolvedUrl;
			if (resolvedUrl === undefined) {
				throw new Error("Resolved Url not available on attached container");
			}
			const url = await container.getAbsoluteUrl("/");
			if (url === undefined) {
				throw new Error("Container has no absolute url");
			}

			return {
				getSharingUrl: url,
				getItemId: resolvedUrl.itemId,
				getDriveId: resolvedUrl.driveId,
			};
		};

		return {
			getTenantAttributes: getAttributes,
			audience: new OdspAudience(container),
		};
	}
}
