import ISessionStorage from "./ISessionStorageService";
import { ISessionStorageElement } from "../model/ISessiongStorageElement";
import { PnPClientStorage } from "@pnp/common";



export default class SessionStorage implements ISessionStorage {

  setSotareElementByKey(key: string, value: any, minutesToExpired: number): void {
    const storage = new PnPClientStorage();
    let currentMin = new Date().getMinutes();
    storage.session.put(key, value, new Date(new Date().setMinutes(currentMin + minutesToExpired)));
  }


  getSotareElementByKey(key: string): ISessionStorageElement {
    const storage = new PnPClientStorage();
    const sessionElement = storage.session.get(key);
    const currentTimeStamp = new Date().getTime();
    if (sessionElement) {
      return { hasExpired: currentTimeStamp > sessionElement.expirationTime, elements: sessionElement.value }

    }
    else {
      return { hasExpired: true, elements: null }

    }

  }

}
