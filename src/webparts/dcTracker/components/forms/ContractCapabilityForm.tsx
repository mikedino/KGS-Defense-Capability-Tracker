import * as React from "react";
import { Callout, DirectionalHint, Stack, Text, TextField } from "@fluentui/react";
import { ICapabilityItem } from "../common/props";
import { DataSource } from "../data/ds";
import styles from "../Dct.module.scss";

export interface IContractCapabilityFormProps {
    capabilities: ICapabilityItem[];
    linkedCapabilityIds: Set<number>;
    selectedCapability?: ICapabilityItem;
    summary: string;
    isEditMode?: boolean;
    onCapabilityChange: (capability?: ICapabilityItem) => void;
    onSummaryChange: (summary: string) => void;
}

const maxCapabilityResults = 30;

export const ContractCapabilityForm: React.FC<IContractCapabilityFormProps> = ({
    capabilities,
    linkedCapabilityIds,
    selectedCapability,
    summary,
    isEditMode = false,
    onCapabilityChange,
    onSummaryChange
}) => {
    const targetRef = React.useRef<HTMLDivElement>(null);
    const [searchText, setSearchText] = React.useState<string>(selectedCapability?.Title ?? "");
    const [searchQuery, setSearchQuery] = React.useState<string>("");
    const [isOpen, setIsOpen] = React.useState<boolean>(false);
    const [activeIndex, setActiveIndex] = React.useState<number>(0);

    React.useEffect(() => {
        setSearchText(selectedCapability?.Title ?? "");
    }, [selectedCapability?.Id]);

    const filteredCapabilities = React.useMemo<ICapabilityItem[]>(() => {
        const search = searchQuery.trim().toLowerCase();

        return capabilities
            .filter((capability) => isEditMode || !linkedCapabilityIds.has(capability.Id))
            .filter((capability) => {
                if (!search) return true;

                const searchableText = [
                    capability.Title,
                    capability.synonyms,
                    DataSource.getConfigText("capabilityType", capability.capabilityTypeTier1),
                    DataSource.getConfigText("capabilityType", capability.capabilityTypeTier2)
                ].filter(Boolean).join(" ").toLowerCase();

                return searchableText.includes(search);
            })
            .sort((a, b) => (a.Title ?? "").localeCompare(b.Title ?? ""))
            .slice(0, maxCapabilityResults);
    }, [capabilities, isEditMode, linkedCapabilityIds, searchQuery]);

    const selectCapability = (capability: ICapabilityItem): void => {
        onCapabilityChange(capability);
        setSearchText(capability.Title ?? "");
        setSearchQuery("");
        setIsOpen(false);
        setActiveIndex(0);
    };

    const resetInput = (): void => {
        setSearchText(selectedCapability?.Title ?? "");
        setSearchQuery("");
        setIsOpen(false);
        setActiveIndex(0);
    };

    return (
        <Stack tokens={{ childrenGap: 14 }}>
            <div ref={targetRef} className={`${styles.formControl} ${styles.stateAutocomplete}`}>
                <TextField
                    label="Capability"
                    required
                    disabled={isEditMode}
                    value={searchText}
                    placeholder="Search by title, synonyms, or capability type..."
                    autoComplete="off"
                    role="combobox"
                    aria-expanded={isOpen}
                    aria-autocomplete="list"
                    onFocus={() => {
                        setSearchQuery("");
                        setIsOpen(true);
                    }}
                    onBlur={resetInput}
                    onChange={(_, value) => {
                        const nextValue = value ?? "";
                        setSearchText(nextValue);
                        setSearchQuery(nextValue);
                        setIsOpen(true);
                        setActiveIndex(0);
                        if (selectedCapability && nextValue !== selectedCapability.Title) {
                            onCapabilityChange(undefined);
                        }
                    }}
                    onKeyDown={(event) => {
                        if (event.key === "ArrowDown") {
                            event.preventDefault();
                            setIsOpen(true);
                            setActiveIndex((index) => Math.min(index + 1, Math.max(filteredCapabilities.length - 1, 0)));
                        } else if (event.key === "ArrowUp") {
                            event.preventDefault();
                            setActiveIndex((index) => Math.max(index - 1, 0));
                        } else if (event.key === "Enter" && isOpen && filteredCapabilities[activeIndex]) {
                            event.preventDefault();
                            selectCapability(filteredCapabilities[activeIndex]);
                        } else if (event.key === "Escape") {
                            event.preventDefault();
                            resetInput();
                        }
                    }}
                />

                {isOpen && !isEditMode && (
                    <Callout
                        target={targetRef.current}
                        directionalHint={DirectionalHint.bottomLeftEdge}
                        gapSpace={2}
                        isBeakVisible={false}
                        setInitialFocus={false}
                        styles={{ root: { width: targetRef.current?.offsetWidth ?? 420 } }}
                    >
                        <div className={styles.stateAutocompleteOptions} role="listbox" aria-label="Capability suggestions">
                            {filteredCapabilities.length ? filteredCapabilities.map((capability, index) => {
                                const typeLabels = [
                                    DataSource.getConfigText("capabilityType", capability.capabilityTypeTier1),
                                    DataSource.getConfigText("capabilityType", capability.capabilityTypeTier2)
                                ].filter(Boolean).join("; ");

                                return (
                                    <button
                                        key={capability.Id}
                                        type="button"
                                        role="option"
                                        aria-selected={index === activeIndex}
                                        className={`${styles.stateAutocompleteOption} ${index === activeIndex ? styles.stateAutocompleteOptionActive : ""}`}
                                        onMouseDown={(event) => event.preventDefault()}
                                        onMouseEnter={() => setActiveIndex(index)}
                                        onClick={() => selectCapability(capability)}
                                    >
                                        <Stack>
                                            <Text>{capability.Title}</Text>
                                            {(capability.synonyms || typeLabels) && (
                                                <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
                                                    {[capability.synonyms, typeLabels].filter(Boolean).join(" • ")}
                                                </Text>
                                            )}
                                        </Stack>
                                    </button>
                                );
                            }) : (
                                <div className={styles.stateAutocompleteEmpty}>No matching capabilities</div>
                            )}
                        </div>
                    </Callout>
                )}
            </div>

            <TextField
                label="Contract Capability Summary"
                className={styles.formControl}
                multiline
                autoAdjustHeight
                rows={5}
                value={summary}
                placeholder="Describe how this capability supports or relates to the contract."
                onChange={(_, value) => onSummaryChange(value ?? "")}
            />
        </Stack>
    );
};
