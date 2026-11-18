import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

export interface INovaPolozkaCommandSetProperties {
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'NovaPolozkaCommandSet';

export default class NovaPolozkaCommandSet extends BaseListViewCommandSet<INovaPolozkaCommandSetProperties> {

  public onInit(): Promise<void> {

    Log.info(LOG_SOURCE, 'Initialized');

    const command: Command | undefined = this.tryGetCommand('COMMAND_1');

    if (command) {
      command.visible = true;
      command.title = 'Nový záznam OTA';
    }

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    switch (event.itemId) {

      case 'COMMAND_1':

        window.open(
          'https://apps.powerapps.com/play/e/16957bf0-1b78-e288-941a-e7e11c01d041/a/9b763c64-607e-4d52-8ae8-8b636cb30f17?tenantId=b233f9e1-5599-4693-9cef-38858fe25406&hint=b6e72755-546c-4afe-88a7-70863138631b&ID=0',
          '_blank'
        );

        break;

      default:
        throw new Error('Unknown command');
    }
  }
}