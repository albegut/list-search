
import { SearchResults } from "@pnp/sp/search";

export default interface ISearchService {
  getSitesStartingWith(tenant: string): Promise<SearchResults>;
  getPathFromResults(results: SearchResults): Array<string>;
}
