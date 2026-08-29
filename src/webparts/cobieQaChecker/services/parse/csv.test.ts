import { detectDelimiter, parseDelimited } from './csv';

describe('parseDelimited', () => {
  it('reads a plain grid', () => {
    expect(parseDelimited('a,b\n1,2', ',')).toEqual([['a', 'b'], ['1', '2']]);
  });

  it('keeps a delimiter that sits inside quotes', () => {
    // COBie list cells (SpaceNames, ComponentNames) are comma-separated *inside*
    // one field. Getting this wrong shifts every later column by one and reports
    // the whole sheet as broken.
    expect(parseDelimited('Name,Spaces\nZ1,"101,102,103"', ','))
      .toEqual([['Name', 'Spaces'], ['Z1', '101,102,103']]);
  });

  it('unescapes doubled quotes', () => {
    expect(parseDelimited('a\n"say ""hi"""', ',')).toEqual([['a'], ['say "hi"']]);
  });

  it('keeps a newline inside a quoted field', () => {
    expect(parseDelimited('a,b\n"one\ntwo",3', ',')).toEqual([['a', 'b'], ['one\ntwo', '3']]);
  });

  it('treats CRLF as one row break', () => {
    expect(parseDelimited('a,b\r\n1,2\r\n', ',')).toEqual([['a', 'b'], ['1', '2']]);
  });

  it('does not append a phantom row for a trailing terminator', () => {
    expect(parseDelimited('a\n1\n', ',')).toEqual([['a'], ['1']]);
  });

  it('keeps a genuinely empty trailing field', () => {
    expect(parseDelimited('a,b\n1,', ',')).toEqual([['a', 'b'], ['1', '']]);
  });
});

describe('detectDelimiter', () => {
  it('picks tab for a TSV header', () => {
    expect(detectDelimiter('Name\tCreatedBy\tCreatedOn')).toBe('\t');
  });

  it('picks comma by default', () => {
    expect(detectDelimiter('Name,CreatedBy,CreatedOn')).toBe(',');
  });

  it('does not pick semicolon when commas separate the header', () => {
    // The regression this pins: a comma-delimited file whose *data* holds
    // semicolon-separated lists must still be read as CSV.
    expect(detectDelimiter('Name,SpaceNames,CreatedBy')).toBe(',');
  });

  it('picks semicolon when it outright wins', () => {
    expect(detectDelimiter('Name;CreatedBy;CreatedOn')).toBe(';');
  });
});
