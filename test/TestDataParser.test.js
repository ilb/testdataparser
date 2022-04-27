import TestDataParser from '../src/TestDataParser';

test('TestDataParser', () => {
  const testDataParser = new TestDataParser('test/interestsstatement.ods');
  const result = testDataParser.parseAllSheets();
  expect(result['var1']['/request/']['begDate']).toEqual(new Date('2020-12-14'));
  // глюк? сдвоенные ячейки (две даты равны в соседних ячейках) во второй ячейке выводятся строкой
  expect(result['var2']['#interestsStatement#'][1]['endDate']).toEqual(new Date('2020-03-01'));
  expect(result['var2']['#interestsStatement#'][2]['endDate']).toEqual(new Date('2020-03-02'));
});
