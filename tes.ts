import { Injectable } from '@angular/core';
import { Router } from '@angular/router';
import { HttpErrorResponse, HttpEvent, HttpHandler, HttpInterceptor, HttpRequest, HttpResponse } from '@angular/common/http';
import * as _ from 'underscore';
import { NotificationService } from '../service/notification.service';
import { Observable, catchError, throwError, finalize } from 'rxjs';

@Injectable()
export class CustomInterceptor implements HttpInterceptor {
  alertCount: number = 0;
  private readonly debouncedSetStatus: () => void;

  constructor(private router: Router, private notificationService: NotificationService) {
    this.debouncedSetStatus = _.debounce(() => this.setStatus(), 10);
  }

  intercept(request: HttpRequest<any>, next: HttpHandler): Observable<HttpEvent<any>> {

    const isPhotoRequest = request.url.includes('/photos/');

    if (!isPhotoRequest) {
      this.setLoadingStatus(true);
    }

    const outgoing = isPhotoRequest
      ? request
      : request.clone({
          headers: request.headers.set('Cache-Control', 'no-cache')
            .set('Pragma', 'no-cache')
        });

    return next.handle(outgoing).pipe(
      catchError((error: HttpErrorResponse) => {
        return this.handleError(error);
      }),
      finalize(() => {
        if (!isPhotoRequest) {
          this.setLoadingStatus(false);
        }
      })
    );
  }

  setLoadingStatus(isLoading: boolean) {
    isLoading ? this.alertCount++ : this.alertCount--;
    this.debouncedSetStatus();
  }

  setStatus() {
    this.notificationService.loadingSubj.next(this.alertCount > 0);
  }

  handleException(event: HttpEvent<any>) {
    if (event instanceof HttpResponse && event.status == 200) {
      let data = event.body;
      if (data && 'success' in data && data.success === false) {
        this.notificationService.setNotificationData('alert-danger', 'HTTP request Failed - ' + data.message, '');
        throwError(data.message);
      }
    }
    return event;
  }

  handleError(err: HttpErrorResponse): Observable<HttpEvent<any>> {
    console.error(err);
    if (err && err.url && err.url.includes('/photos/')) {
      return throwError(() => err);
    }
    if (err) {
      let message = err.error ? (typeof err.error == 'string' ? err.error : err.statusText) : err.message;
      switch (err.status) {
        case 404:
          this.notificationService.setNotificationData('alert-danger', 'HTTP Request Error-' + err.status + ' Not Found', '');
          break;
        case 403:
          break;
        case 400:
          this.notificationService.setNotificationData('alert-danger', 'HTTP Request Error-' + err.status, '');
          break;
        case 302:
          console.error('err1');
          console.error(err);
          this.notificationService.setNotificationData('alert-danger', 'Session expired, please retry or refresh the page ' + err, '');
          break;
        case 0:
          console.error('err2');
          console.error(err);
          this.notificationService.setNotificationData('alert-danger', 'Login failed.' + message, '');
          break;
        default:
          message = (err.error ? (typeof err.error == 'string' ? err.error : err.statusText) : err.message) || 'Error occured while trying to connect ' + (err.url && err.url.slice(err.url.indexOf('rest') + 4) || err.error.target.__zone_symbol_xhrURL.slice(6));
          this.notificationService.setNotificationData('alert-danger', 'HTTP Request Error-' + err.status + ':' + message, '');
      }
      return throwError(() => err);
    }
    return err;
  }
}
