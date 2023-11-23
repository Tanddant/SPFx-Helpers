import { HttpRequestError } from "@pnp/queryable";

export namespace ErrorHelper {
    export const PnPjsErrorWrapper = async (func: () => Promise<any>) => {
        try {
            return await func();
        } catch (e) {
            if (e?.isHttpRequestError) {
                const json = await (<HttpRequestError>e).response.json();
                alert(typeof json["odata.error"] === "object" ? json["odata.error"].message.value : e.message);

            } else {
                alert(e);
            }
            throw e
        }
    }
}