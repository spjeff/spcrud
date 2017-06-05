/* tslint:disable:no-unused-variable */

import { TestBed, async, inject } from '@angular/core/testing';
import { Spcrud } from './spcrud';

describe('Spcrud', () => {
  beforeEach(() => {
    TestBed.configureTestingModule({
      providers: [Spcrud]
    });
  });

  it('should ...', inject([Spcrud], (service: Spcrud) => {
    expect(service).toBeTruthy();
  }));
});
