import { TimeConversionPipe } from './timeconversion.pipe';

describe('TimeConversionPipe', () => {
  it('create an instance', () => {
    const pipe = new TimeConversionPipe();
    expect(pipe).toBeTruthy();
  });
});
