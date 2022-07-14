import { HumanFileSizePipe } from './human-file-size.pipe';

describe('HumanFileSizePipe', () => {
  it('create an instance', () => {
    const pipe = new HumanFileSizePipe();
    expect(pipe).toBeTruthy();
  });
});
