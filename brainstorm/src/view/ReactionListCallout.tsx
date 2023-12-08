import { Icon, Label, Stack } from "@fluentui/react";

import React from "react";
import { PersonaList } from "./PersonaList";
import { OdspMember } from "../odsp-client";

export type ReactionListCalloutProps = {
	label: string;
	usersToDisplay: OdspMember[];
	reactionIconName?: string;
};

export function ReactionListCallout(props: ReactionListCalloutProps) {
	return (
		<div>
			<Stack horizontal tokens={{ childrenGap: 10 }}>
				{props.reactionIconName && (
					<Icon
						iconName={props.reactionIconName}
						style={{ fontSize: 15, alignSelf: "center" }}
					></Icon>
				)}
				<Label>Like Reactions</Label>
			</Stack>
			<PersonaList users={props.usersToDisplay} />
		</div>
	);
}
