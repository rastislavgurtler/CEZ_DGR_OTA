import { Log } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';

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

  private _powerAppUrl: string = '';
  private _aktivni: boolean = false;
  private _listGuid: string = '';

  public async onInit(): Promise<void> {

    Log.info(LOG_SOURCE, 'Initialized');

    const command: Command | undefined = this.tryGetCommand('COMMAND_1');

    if (command) {
      command.title = 'Nový záznam';
      command.visible = false;
    }

    try {

      // PowerAppUrl
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('list_Config-C')/items?$filter=Title eq 'PowerAppUrl'`,
        SPHttpClient.configurations.v1
      );

      const data = await response.json();

      if (data.value.length > 0) {
        this._powerAppUrl = data.value[0].Hodnota;
      }

      // Aktivni
      const activeResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('list_Config-C')/items?$filter=Title eq 'Aktivni'`,
        SPHttpClient.configurations.v1
      );

      const activeData = await activeResponse.json();

      if (activeData.value.length > 0) {
        this._aktivni =
          activeData.value[0].Hodnota === 'Ano';
      }

      // ListGuid
      const listGuidResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('list_Config-C')/items?$filter=Title eq 'ListGuid'`,
        SPHttpClient.configurations.v1
      );

      const listGuidData = await listGuidResponse.json();

      if (listGuidData.value.length > 0) {
        this._listGuid = listGuidData.value[0].Hodnota;
      }

      if (command) {

        const currentListId =
          this.context.pageContext.list?.id
            .toString()
            .replace('{', '')
            .replace('}', '')
            .toLowerCase();

        const configListId =
          this._listGuid
            .replace('{', '')
            .replace('}', '')
            .toLowerCase();

        command.visible =
          this._aktivni &&
          currentListId === configListId;

      }

      this.raiseOnChange();

    } catch (error) {

      console.error(
        'Chyba při načítání konfigurace',
        error
      );

    }

    return Promise.resolve();
  }

  public onExecute(
    event: IListViewCommandSetExecuteEventParameters
  ): void {

    switch (event.itemId) {

      case 'COMMAND_1':

        if (this._powerAppUrl) {

          window.open(
            this._powerAppUrl,
            '_blank'
          );

        }

        break;

      default:
        throw new Error('Unknown command');
    }
  }
}