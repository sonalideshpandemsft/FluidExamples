import { IPersonaStyles, List, Persona, PersonaSize } from "@fluentui/react";

import React from "react";
import { OdspMember } from "../odsp-client";

export function PersonaList(props: { users: OdspMember[] }) {
	const personaStyles: Partial<IPersonaStyles> = {
		root: {
			marginTop: 10,
		},
	};

	const renderPersonaListItem = (item?: OdspMember) => {
		return (
			item && (
				<Persona
					text={item.name}
					size={PersonaSize.size24}
					styles={personaStyles}
				></Persona>
			)
		);
	};
	return <List items={props.users} onRenderCell={renderPersonaListItem}></List>;
}
