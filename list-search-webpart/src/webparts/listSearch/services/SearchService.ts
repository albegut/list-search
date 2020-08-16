import ISearchService from "./ISearchService";
import { sp } from "@pnp/sp";
import "@pnp/sp/search";
import { SearchQueryBuilder, SearchResults, ISearchQuery } from "@pnp/sp/search";

export default class SearchService implements ISearchService {

  public async getSitesStartingWith(tenantUrl: string): Promise<SearchResults> {
    try {
      const _searchQuerySettings: ISearchQuery = {
        TrimDuplicates: true,
        RowLimit: 500,
        SelectProperties: ["Path"],

      }
      let query = SearchQueryBuilder(`Path:${tenantUrl}/* AND contentclass:STS_Site`, _searchQuerySettings)
      return sp.search(query);
    } catch (error) {
      return Promise.reject(error);
    }
  }

  public getPathFromResults(results: SearchResults): Array<string> {
    return this.getPropertyFromResults(results, "Path");
  }

  private getPropertyFromResults(results: SearchResults, property: string): Array<string> {
    let urls: Array<string> = [];

    results.PrimarySearchResults.map(result => {
      if (urls.indexOf(result[property]) < 0) {
        urls.push(result[property]);
      }
    })
    return urls;
  }
}
