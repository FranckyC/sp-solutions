import { Context, HttpMethod, HttpRequest, HttpResponse, HttpStatusCode } from 'azure-functions-ts-essentials';
import * as fs from 'fs';
import * as path from 'path';
import { IFunctionLocalSettings } from './config/IFunctionLocalSettings';

/*
 * Routes the request to the default controller using the relevant method.
 */
export async function run(context: Context, req: HttpRequest): Promise<HttpResponse> {

  let res: HttpResponse;

  switch (req.method) {
    case HttpMethod.Get:
        break;

    case HttpMethod.Post:

      // The user raw query from the search box
      const parameter1 = req.body
      ? req.body.parameter1
      : undefined;

      // Using a dedicated JSON settings file per function allows you to have multiple configurations for different environments
      // Then depending on your local.settings.json (shared by all functions), you just have to set the correct file path
      const functionLocalSettings = process.env['DemoFunction_LocalSettingsFilePath'];

      // Get Function settings
      const settingsFile = fs.readFileSync(path.resolve(__dirname, `./${functionLocalSettings}`), { encoding: 'utf8' });
      const settingsData = JSON.parse(settingsFile.toString()) as IFunctionLocalSettings;
      const mySettingValue = settingsData.mySetting;

      try {

        if (!parameter1) {
          throw new Error("'parameter1' is not defined");
        }

        // Response object passed in body
        const response = {
          value: `You send 'parameter1' with value '${parameter1}' and function setting ${mySettingValue}`
        }

        res = {
          status: HttpStatusCode.OK,
          body: response
        };

      } catch (error) {
        res = {
          status: HttpStatusCode.InternalServerError,
          body: {
            error: {
              type: 'function_error',
              message: error.message
            }
          }
        };
      }

      break;
    case HttpMethod.Patch:
      break;
    case HttpMethod.Delete:
      break;

    default:
      res = {
              status: HttpStatusCode.MethodNotAllowed,
              body: {
                error: {
                  type: 'not_supported',
                  message: `Method ${req.method} not supported.`
                }
              }
            };
  }

  return res;
}
