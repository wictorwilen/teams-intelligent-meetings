import * as ACData from "adaptivecards-templating";

const cardTemplate = require("./meetingSummary.json");
const MeetingSummary = (data: any) => {
    const template = new ACData.Template(cardTemplate);
    return template.expand({
        $root: data
    });
};

export default MeetingSummary;
