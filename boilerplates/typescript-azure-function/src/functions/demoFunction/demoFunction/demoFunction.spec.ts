import { Context, HttpMethod, HttpRequest, HttpStatusCode } from 'azure-functions-ts-essentials';
import { run } from '../demoFunction';
import * as $ from '../../../../tools/build/helpers';
import * as fs from 'fs';

// Build the process.env object according to Azure Function Application Settings
const localSettingsFile = $.root('./src/local.settings.json');
const settings = JSON.parse(fs.readFileSync(localSettingsFile, { encoding: 'utf8' })
                    .toString()).Values;

// Get all Azure Function settings so we can use them in tests
Object.keys(settings)
    .map(key => {
        process.env[key] = settings[key];
});

describe('POST /api/demoFunction', () => {

    it('should throw a warning message if the \'parameter1\' body parameter is empty', async () => {

        const mockContext: Context = {
            done: (err, response) => {
                expect(err).toBeUndefined();
                expect(response.status).toEqual(HttpStatusCode.InternalServerError);
                expect(response.body.error.message).toBeDefined();
            }
          };

        const mockRequest: HttpRequest = {
            method: HttpMethod.Post,
            headers: { 'content-type': 'application/json' },
            body: {
                parameter1: ''
            }
        };

        try {
            await run(mockContext, mockRequest);
        } catch (e) {
            fail(e);
        }
    });
});
