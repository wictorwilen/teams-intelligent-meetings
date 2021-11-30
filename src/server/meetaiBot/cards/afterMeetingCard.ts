import * as ACData from "adaptivecards-templating";
import meetings from "../../meetingManager";
import * as incidents from "../../incidents.json";

const cardTemplate = require("./afterMeetingCard.json");
const AfterMeetingCard = (meetingId: string) => {
    const meeting = meetings.getById(meetingId);
    const incident = incidents.find(i => i.id === meeting!.data.incident);
    const template = new ACData.Template(cardTemplate);
    return template.expand({
        $root: { ...incident, ...meeting }
    });
};

export default AfterMeetingCard;
