import TestDataParser from '../src/TestDataParser';

test('TestDataParser', () => {
  const testDataParser = new TestDataParser('test/interestsstatement.ods');
  const result = testDataParser.parseAllSheets();
  expect(result.get('var1').get('/request/').get('from')).toEqual(new Date('2020-12-14'));
});
