import * as React from "react";
import { Text, Persona, PersonaSize, TooltipHost } from "@fluentui/react";
import { IPeoplePickerExtended } from "../data/props";

export interface PeoplePersonaProps {
  person?: IPeoplePickerExtended;
  size?: PersonaSize;
  fallbackText?: string;
  showDetails?: boolean; // toggle name + job title display
}

const getUserPhotoUrl = (email: string): string => `/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`;


/**
 * People Persona display helper
 */
export const PeoplePersona: React.FC<PeoplePersonaProps> = ({
  person,
  size = PersonaSize.size40,
  fallbackText = "No user set",
  showDetails = false
}) => {

  if (!person) {
    return (
      <Text styles={{ root: { color: "gray", fontStyle: "italic" } }}>
        {fallbackText}
      </Text>
    );
  }

  return (
    <TooltipHost
      content={
        <div>
          <strong>{person.Title}</strong>
          {person.JobTitle && (
            <>
              <br />
              {person.JobTitle}
            </>
          )}
          <br />
          {person.EMail}
        </div>
      }
      calloutProps={{ gapSpace: 0 }}
      styles={{ root: { display: "inline-block" } }}
    >
      <Persona
        text={person.Title}
        secondaryText={person.JobTitle}
        size={size}
        imageUrl={person.EMail ? getUserPhotoUrl(person.EMail) : undefined}
        hidePersonaDetails={!showDetails}
      />
    </TooltipHost>
  );
};