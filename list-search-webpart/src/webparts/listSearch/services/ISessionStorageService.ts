import { ISessionStorageElement } from "../model/ISessiongStorageElement";

export default interface ISessionStorage {
  setSotareElementByKey(key: string, value: any, minutesToExpired: number): void;
  getSotareElementByKey(key: string): ISessionStorageElement;
}
