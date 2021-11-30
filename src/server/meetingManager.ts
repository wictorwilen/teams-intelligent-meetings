import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta";
import { JsonDB } from "node-json-db";
import { Config } from "node-json-db/dist/lib/JsonDBConfig";

const _storage = new JsonDB(new Config("meetings.db.json", true, false, "/"));

declare type Participant = {
    id: string;
    joinedTimes: number;
    handRaised: boolean;
};

declare type Meeting = {
    id?: string | undefined;
    recording?: boolean | undefined;
    threadId: string;
    participants: Participant[];
    data: {
        incident: string;
    };
    activeParticipants: number;
};

// INFO: sorry - this is a hack!
class MeetingManager {
    // INFO: not multi-tenant as of now
    private Meetings: Meeting[];

    constructor() {
        if (_storage.exists("/meetings")) {
            this.Meetings = _storage.getData("/meetings");
        } else {
            this.Meetings = [];
            _storage.push("/meetings", []);
        }
    }

    public getById = (id: string): Meeting | undefined => {
        const meeting = this.Meetings.find(m => {
            return m.id === id;
        });
        return meeting;
    }

    public getByThreadId = (threadId: string): Meeting | undefined => {
        const meeting = this.Meetings.find(m => {
            return m.threadId === threadId;
        });
        return meeting;
    }

    public add = (meeting: Meeting): void => {
        this.Meetings.push(meeting);
        _storage.push("/meetings", this.Meetings, true);
    }

    public delete = (id: string): void => {
        this.Meetings = this.Meetings.filter(m => m.id !== id);
        _storage.push("/meetings", this.Meetings, true);
    }

    public update = (meeting: Meeting): void => {
        if (meeting.id) {
            this.delete(meeting.id);
        } else {
            this.Meetings = this.Meetings.filter(m => m.threadId !== meeting.threadId);
        }
        this.add(meeting);
    }

    public updateParticipant = (meeting: Meeting, participant: MicrosoftGraphBeta.Participant): void => {
        if (participant.info && participant.info.identity) {
            const part = meeting.participants.find(p =>
                (participant.info?.identity?.user &&
                    p.id === participant.info.identity.user.id) ||
                (participant.info?.identity?.application &&
                    p.id === participant.info.identity.application.id
                ));

            if (!part) {
                meeting.participants.push({
                    id: participant.info.identity.user ? participant.info.identity.user.id : participant.info.identity.application?.id
                } as any);
            }
        }
        this.update(meeting);
    }

}

export default new MeetingManager();
