import { ITag } from "office-ui-fabric-react";

export default class SuggestionTag implements ITag {
    name: string;
    key: string | number;
    info: string;
}