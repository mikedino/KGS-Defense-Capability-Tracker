import * as React from "react";
import { useState, useMemo } from "react";
import { DetailsList, IDetailsListProps } from '@fluentui/react/lib/DetailsList';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Stack, IconButton, Text } from '@fluentui/react';

interface PaginatedDetailsListProps extends IDetailsListProps {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    items: any[]; // override to control slice
    pageSizeOptions?: number[];
    defaultPageSizeOption?: number;
    showFirstLastButtons?: boolean;
}

const PaginatedDetailsList: React.FC<PaginatedDetailsListProps> = ({
    items,
    pageSizeOptions = [5, 10, 20, 50],
    defaultPageSizeOption = 10,
    showFirstLastButtons = true,
    ...detailsListProps
}) => {
    const [currentPage, setCurrentPage] = useState(0);
    const [pageSize, setPageSize] = useState(defaultPageSizeOption);

    const totalPages = Math.ceil(items.length / pageSize);

    const pagedItems = useMemo(() => {
        const start = currentPage * pageSize;
        return items.slice(start, start + pageSize);
    }, [items, currentPage, pageSize]);

    const changePageSize = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
        if (option) {
            setPageSize(option.key as number);
            setCurrentPage(0);
        }
    };

    const goToPage = (index: number): void => {
        if (index >= 0 && index < totalPages) {
            setCurrentPage(index);
        }
    };

    const smallFontStyle = { fontSize: 12 };

    return (
        <Stack tokens={{ childrenGap: 10 }}>
            <DetailsList
                items={pagedItems}
                {...detailsListProps}
            />

            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Stack horizontal verticalAlign="center">
                    {showFirstLastButtons && (
                        <IconButton
                            iconProps={{ iconName: 'DoubleChevronLeft', title: 'First Page' }}
                            onClick={() => goToPage(0)}
                            disabled={currentPage === 0}
                            styles={{ root: smallFontStyle, icon: smallFontStyle }}
                        />
                    )}
                    <IconButton
                        iconProps={{ iconName: 'ChevronLeft', title: 'Previous Page' }}
                        onClick={() => goToPage(currentPage - 1)}
                        disabled={currentPage === 0}
                        styles={{ root: smallFontStyle, icon: smallFontStyle }}
                    />
                    <Text styles={{ root: {...smallFontStyle, padding: '0 10px'} }}>
                        Page {currentPage + 1} of {totalPages}
                    </Text>
                    <IconButton
                        iconProps={{ iconName: 'ChevronRight', title: 'Next Page' }}
                        onClick={() => goToPage(currentPage + 1)}
                        disabled={currentPage >= totalPages - 1}
                        styles={{ root: smallFontStyle, icon: smallFontStyle }}
                    />
                    {showFirstLastButtons && (
                        <IconButton
                            iconProps={{ iconName: 'DoubleChevronRight', title: 'Last Page' }}
                            onClick={() => goToPage(totalPages - 1)}
                            disabled={currentPage >= totalPages - 1}
                            styles={{ root: smallFontStyle, icon: smallFontStyle }}
                        />
                    )}
                </Stack>

                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <Text styles={{ root: smallFontStyle }}>Page Size:</Text>
                    <Dropdown
                        selectedKey={pageSize}
                        options={pageSizeOptions.map(n => ({ key: n, text: `${n} per page` }))}
                        onChange={changePageSize}
                        styles={{
                            dropdown: { ...smallFontStyle, width: 120 },
                            title: smallFontStyle,
                            dropdownItem: smallFontStyle
                        }}
                    />
                </Stack>
            </Stack>
        </Stack>
    );

};

export default PaginatedDetailsList;
