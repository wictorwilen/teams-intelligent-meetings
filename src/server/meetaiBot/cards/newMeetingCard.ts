import * as ACData from "adaptivecards-templating";
const cardTemplate = require("./newMeetingCard.json");
const NewMeetingCard = (data: any) => {
    const template = new ACData.Template(cardTemplate);
    return template.expand({
        $root: data
    });
};

export default NewMeetingCard;
