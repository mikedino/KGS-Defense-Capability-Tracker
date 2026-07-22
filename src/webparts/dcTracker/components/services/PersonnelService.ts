import { Types, Web } from "gd-sprest-bs";
import { formatError } from "../common/utils";

export class PersonnelService {

    static addUser(groupName: string, userId: number): Promise<Types.SP.User> {
        return new Promise<Types.SP.User>((resolve, reject) => {
            Web().SiteGroups().getByName(groupName).Users().addUserById(userId).execute(
                //success
                resp => {
                    if (resp) {
                        console.info(`Added new user!`)
                        resolve(resp);
                    }
                },
                //error
                (error) => {
                    const err = formatError(error);
                    console.error(`Error adding user ${err}`);                    
                    reject(error);
                }
            )
        })
    }

    static delete(groupName: string, userId: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Web().SiteGroups().getByName(groupName).Users().removeById(userId).execute(
                //success
                () => {
                    console.info(`Deleted User!`)
                    resolve();
                },
                //error
                (error) => {
                    const err = formatError(error);
                    console.error(`Error removing user: ${err}`);
                    reject(error);
                }
            )
        })
    }

}