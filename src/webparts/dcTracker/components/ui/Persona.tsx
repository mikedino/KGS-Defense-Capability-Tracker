import * as React from "react";
import { Text, Persona, PersonaSize, TooltipHost } from "@fluentui/react";
import { IPeoplePickerExtended } from "../common/props";

export interface PeoplePersonaProps {
  person?: IPeoplePickerExtended;
  size?: PersonaSize;
  fallbackText?: string;
  showDetails?: boolean; // toggle name + job title display
}

const getUserPhotoUrl = (email: string): string => `/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`;

const getInitial = (value?: string): string => (value ?? "").trim().charAt(0).toUpperCase();

const getPersonInitials = (displayName?: string): string | undefined => {
  const name = (displayName ?? "").trim();
  if (!name) return undefined;

  if (name.includes(",")) {
    const [lastName, givenNames] = name.split(",", 2);
    const firstName = givenNames.trim().split(/\s+/)[0];
    return `${getInitial(firstName)}${getInitial(lastName)}` || undefined;
  }

  const parts = name.split(/\s+/).filter(Boolean);
  return `${getInitial(parts[0])}${getInitial(parts.length > 1 ? parts[parts.length - 1] : "")}` || undefined;
};


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
        imageInitials={getPersonInitials(person.Title)}
        hidePersonaDetails={!showDetails}
      />
    </TooltipHost>
  );
};
