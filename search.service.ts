import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { URL_SEARCH, URL_SEARCH_EX } from '../data/constants';
import { EMPTY, map, Observable, of } from 'rxjs';
import { UserCod } from '../data/user-cod';
import { UserGenpop } from '../data/user-genpop';
import { QuickSearchResult } from '../data/quick-search-result';

const cityCollator = new Intl.Collator('en', { sensitivity: 'base' });

function normalize(value?: string): string {
  return (value ?? '').trim().toLowerCase();
}

function matchRank(user: UserCod | UserGenpop, keyword: string): number {
  const fullName = normalize(user.FullName);
  const accounts = [normalize(user.LoginID), normalize(user.MSID)];
  if (fullName === keyword || accounts.includes(keyword)) return 0;

  const names = [normalize(user.FirstName), normalize(user.LastName), ...fullName.split(/[\s,]+/)];
  if (names.includes(keyword)) return 1;
  if ([fullName, ...names, ...accounts].some(value => value.startsWith(keyword))) return 2;
  return 3;
}

function compareCity(left: UserCod | UserGenpop, right: UserCod | UserGenpop): number {
  const a = left.City?.trim() || '';
  const b = right.City?.trim() || '';
  if (!a) return b ? 1 : 0;
  if (!b) return -1;
  return cityCollator.compare(a, b);
}

@Injectable({
  providedIn: 'root'
})
export class SearchService {

  constructor(private httpClient: HttpClient) { }

  quickSearch(query: string): Observable<(UserCod | UserGenpop)[]> {
    if (!query || query.length < 2) {
      return of([]);
    }
    return this.httpClient.get<(UserCod | UserGenpop)[]>(
      URL_SEARCH, {
        params: { Keyword: query }
      });
  }

  quickSearchEx(query: string, enableSizeLimit = true, peopleOrder: 'city' | 'relevance' = 'city'): Observable<QuickSearchResult> {
    if (!query || query.length < 2) {
      return EMPTY;
    }
    const keyword = normalize(query);
    const comparePeople = peopleOrder === 'relevance' && keyword
      ? (left: UserCod | UserGenpop, right: UserCod | UserGenpop) =>
          matchRank(left, keyword) - matchRank(right, keyword) || compareCity(left, right)
      : compareCity;
    return this.httpClient.get<QuickSearchResult>(
      URL_SEARCH_EX, {
        params: {
          Keyword: query,
          EnableSizeLimit: enableSizeLimit }
      }).pipe(
        map(result => ({
          ...result,
          Users: [...(result?.Users ?? [])].sort(comparePeople)
        }))
      );
  }
}
