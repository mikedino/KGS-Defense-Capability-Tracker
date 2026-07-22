export const formatError = (error: unknown): string => {

    // if it's a string, return the string
    if (typeof error === "string") {
        return error;
    }

    if (error instanceof Error) {
        // If it's an instance of Error, use its message
        return error.message;
    } else if (typeof error === 'object' && error !== null) {
        // Check if the error has a 'response' key
        if ('response' in error) {
            const response = (error as { response: string }).response; // Type assertion
            try {
                // Parse the JSON string in the 'response' key
                const parsedResponse = JSON.parse(response);
                // Check if the parsed response has the expected structure
                if (parsedResponse.error && parsedResponse.error.message) {
                    return parsedResponse.error.message.value || "An unknown error occurred.";
                }
            } catch (e) {
                // If parsing fails, return a generic error message
                return `An error occurred while processing the error response. ${JSON.stringify(e)}`;
            }
        }
        // Fallback to extracting a message from the error object
        return (error as { message?: string }).message || JSON.stringify(error);
    } else {
        // Fallback for any other type
        return "An unknown error occurred.";
    }
};

/**
 * Helper function to format currency.
 * @param amount - The number or string to be formatted.
 * @param locale - The locale string (default is 'en-US').
 * @param currency - The currency code (default is 'USD').
 * @returns A formatted currency string or an empty string if the input is invalid.
 */
export const formatCurrency = (
    amount: number | string,
    locale: string = 'en-US',
    currency: string = 'USD'
): string => {
    const number = parseFloat(amount as string);
    return isNaN(number)
        ? ''
        : new Intl.NumberFormat(locale, {
            style: 'currency',
            currency,
            minimumFractionDigits: 0, // Ensures no cents are displayed
            maximumFractionDigits: 0, // Ensures no cents are displayed
        }).format(number);
};


/**
 * Helper function to format dates.
 * @param dateInput - The date input (string, number, or Date object).
 * @param locale - The locale string (default is 'en-US').
 * @returns A formatted date string or an empty string if the input is invalid.
 */
export const formatDate = (
    dateInput: string,
    locale: string = 'en-US'
): string => {
    if (dateInput === undefined || dateInput === null) {
        return '-'; // Handle undefined, empty string, or falsy inputs
    }

    const date = new Date(dateInput);
    return isNaN(date.getTime()) ? '-' : date.toLocaleDateString(locale);
};

/**
 * Helper function to encode list name with underscore or space or other characters that need encoding.
 * @param listName - The list name input (string).
 * @returns List name encoded
 */
export const encodeListName = (
    listName: string
): string => {
    return listName.replace(/[^a-zA-Z0-9]/g, (char) => {
        const hex = char.charCodeAt(0).toString(16).padStart(4, "0");
        return `_x${hex}_`;
    });
}