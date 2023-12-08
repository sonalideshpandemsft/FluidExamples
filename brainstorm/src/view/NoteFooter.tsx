import React from "react";
import { Text } from "@fluentui/react";

import { NoteData } from "../Types";
import { OdspMember } from "../odsp-client";

export type NoteFooterProps = { currentUser: OdspMember } & Pick<NoteData, "lastEdited">;

//deplay time in ms for waiting note content changes to be settle
const delay = 2000;

export function NoteFooter(props: NoteFooterProps) {
	const { currentUser, lastEdited } = props;
	let lastEditedMemberName;

	// Only display the author name if 2 seconds have elapsed since the note was last edited.
	if (Date.now() - lastEdited.time >= delay) {
		// Dynamically display the last edited author name based on if the user is the last edited author
		// If the user is the last edited author, display "you", otherwise, display the author's name.
		lastEditedMemberName =
			currentUser?.name === lastEdited.userName ? "you" : lastEdited.userName;
	} else {
		// Display "..." to indicate the note is being edited.
		lastEditedMemberName = "...";
	}

	return (
		<div style={{ flex: 1 }}>
			<Text
				block={true}
				nowrap={true}
				variant={"medium"}
				styles={{
					root: { alignSelf: "center", marginLeft: 7 },
				}}
			>
				Last edited by {lastEditedMemberName}
			</Text>
		</div>
	);
}
