import { provideHttpClient } from '@angular/common/http';
import { HttpTestingController, provideHttpClientTesting } from '@angular/common/http/testing';
import { TestBed } from '@angular/core/testing';
import { firstValueFrom } from 'rxjs';
import { FilterUsersPipe } from '../pipes/filter-users.pipe';
import { SearchService } from './search.service';

describe('People search ordering', () => {
  let service: SearchService;
  let http: HttpTestingController;

  beforeEach(() => {
    TestBed.configureTestingModule({ providers: [provideHttpClient(), provideHttpClientTesting()] });
    service = TestBed.inject(SearchService);
    http = TestBed.inject(HttpTestingController);
  });

  afterEach(() => http.verify());

  it('sorts across page boundaries before the People table takes its page', async () => {
    const response = firstValueFrom(service.quickSearchEx('lee', false));
    const request = http.expectOne(req => req.params.get('EnableSizeLimit') === 'false');
    const london = Array.from({ length: 198 }, () => ({ City: 'London' }));
    request.flush({ Users: [{ FullName: 'Lee', City: 'New York' }, ...london, { City: 'Baltimore' }] });

    const users = (await response).Users!;
    const filter = new FilterUsersPipe();
    expect(users.length).toBe(200);
    expect(filter.transform(users, {}, 1, 15)[0].City).toBe('Baltimore');
    expect(filter.transform(users, {}, 14, 15).at(-1)?.City).toBe('New York');
  });

  it('keeps same-city matches stable and missing cities last in the dropdown', async () => {
    const response = firstValueFrom(service.quickSearchEx('lee', true, 'relevance'));
    const request = http.expectOne(req => req.params.get('EnableSizeLimit') === 'true');
    request.flush({
      Users: [
        { MSID: 'blank', City: ' ' }, { MSID: 'one', City: ' Bengaluru ' },
        { MSID: 'calgary', City: 'Calgary' }, { MSID: 'two', City: 'bengaluru' },
        { MSID: 'unknown' }, { MSID: 'baltimore', City: 'Baltimore' }
      ]
    });
    expect((await response).Users?.map(user => user.MSID))
      .toEqual(['baltimore', 'one', 'two', 'calgary', 'blank', 'unknown']);
  });

  it('ranks exact identities, complete name parts, prefixes and other matches, then city', async () => {
    const query = '  StEvEn  ';
    const response = firstValueFrom(service.quickSearchEx(query, true, 'relevance'));
    const request = http.expectOne(req =>
      req.params.get('Keyword') === query && req.params.get('EnableSizeLimit') === 'true');
    request.flush({
      Users: [
        { MSID: 'other', FullName: 'Asteven', City: 'Athens' },
        { MSID: 'prefix', FullName: 'Alice Stevens', City: 'Bengaluru' },
        { MSID: 'first', FirstName: 'Steven', City: 'Shanghai' },
        { MSID: 'last', FullName: 'Steven, Alice', City: 'London' },
        { MSID: 'account', LoginID: 'STEVEN', City: 'Zurich' },
        { MSID: 'steven', City: 'New York' },
        { MSID: 'full', FullName: 'Steven', City: 'Berlin' }
      ],
      MailGroups: [{ MailGroup: { Name: 'Z' } }, { MailGroup: { Name: 'A' } }]
    });

    const result = await response;
    expect(result.Users?.map(user => user.MSID))
      .toEqual(['full', 'steven', 'account', 'last', 'first', 'prefix', 'other']);
    expect(result.MailGroups?.map(group => group.MailGroup.Name)).toEqual(['Z', 'A']);
  });
});
