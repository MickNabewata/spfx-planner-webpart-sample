import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PlannerWebpartWebPart.module.scss';
import * as strings from 'PlannerWebpartWebPartStrings';

// Microsoft Graphへの問い合わせ実行のために追加 パッケージ追加は不要
import { MSGraphClient } from '@microsoft/sp-http';
import { GraphError } from '@microsoft/microsoft-graph-client';

// Microsoft Graphとのやり取りに使う型があるほうがコーディングが楽なので追加
// 要パッケージ追加
// npm install @microsoft/microsoft-graph-types --save-dev
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IPlannerWebpartWebPartProps {
  description: string;
}

export default class PlannerWebpartWebPart extends BaseClientSideWebPart<IPlannerWebpartWebPartProps> {

  public render(): void {

    // Plannerタスク一覧を取得
    this.getTasks((error: GraphError, datas: any) => {

      if (error) {
        this.domElement.innerHTML = `
          <div class="${ styles.plannerWebpart}">
            <div class="${ styles.container}">
              <div class="${ styles.error}">${(error) ? JSON.stringify(error) : ''}</div>
            </div>
          </div>
        `;
      }
      else {
        let events: MicrosoftGraph.PlannerTask[] = datas.value;

        // フィルタ処理
        events = events.filter((event) => { 
          // 未完了タスクを取得
          return (event.percentComplete < 100);
        });

        // ソート処理
        events = events.sort((task1, task2) => {
          return (new Date(task1.dueDateTime) > new Date(task2.dueDateTime)) ? 1 : -1;
        });

        this.domElement.innerHTML = `
          <div class="${ styles.plannerWebpart}">
            <div class="${ styles.container}">
              <div>タスク一覧</div>
              ${
                (events && events.length > 0) ?
                  `<table class="${styles.events}">
                    <thead>
                      <tr>
                        <th>プランId</th>
                        <th>バケットId</th>
                        <th>件名</th>
                        <th>進行状況</th>
                        <th>開始日</th>
                        <th>期限</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${events.map((event) => {

                        // 進行状況は数値で表されている
                        let completed: string = '';
                        switch (event.percentComplete) {
                          case 100:
                            completed = '完了済み';
                            break;
                          case 50:
                            completed = '進行中';
                            break;
                          case 0:
                            completed = '開始前';
                            break;
                          default:
                            break;
                        }

                        return `
                          <tr>
                            <td class="${styles.td}">${ event.planId }</td>
                            <td class="${styles.td}">${ event.bucketId }</td>
                            <td class="${styles.td}">${ event.title }</td>
                            <td class="${styles.td}">${ `${event.percentComplete} : ${completed}` }</td>
                            <td class="${styles.td}">${ event.startDateTime }</td>
                            <td class="${styles.td}">${ event.dueDateTime }</td>
                          </tr>
                        `;
                      })}
                    </tbody>
                  </table>
                  <div> 注) プラン名はプランIDを元に別途取得が必要です。リクエスト例) https://graph.microsoft.com/v1.0/planner/plans/FlMlaw_Lzk6_jbT0XznRrsgAE6T2 </div>
                  <div> 注) バケット名はバケットIDを元に別途取得が必要です。リクエスト例) https://graph.microsoft.com/v1.0/planner/buckets/kxtenetJI0y2zMohNoolTMgABBy_ </div>` :
                  'タスクがありません'
              }
            </div>
          </div>
        `;
      }
    });
  }

  /**
    Microsoft Graphからのデータ取得

    Microsoft Graphへの問い合わせにはアクセス許可が必要です。
    そのためconfigフォルダ > package-solution.jsonファイルのsolutionプロパティ内に以下を追記してあります。
    "webApiPermissionRequests": [
      {
        "resource": "Microsoft Graph",
        "scope": "Group.Read.All"
      }
    ]
    また、当パッケージをSharePointのアプリカタログサイトに展開した後、
    SharePoint管理センター > APIの管理 画面で
    当パッケージが要求するアクセス許可(Group.Read.All)の承認が必要です。

    尚、この問い合わせ結果に含まれるプランとバケットに関する情報はIDのみとなっています。
    プラン名やバケット名が必要な場合、別途取得が必要です。
      プラン名リクエスト例) https://graph.microsoft.com/v1.0/planner/plans/FlMlaw_Lzk6_jbT0XznRrsgAE6T2
      バケット名リクエスト例) https://graph.microsoft.com/v1.0/planner/buckets/kxtenetJI0y2zMohNoolTMgABBy_
  */
  protected getTasks(callBack: (error: GraphError, datas: any) => void): Promise<any> {
    return this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        // ユーザー自身(me)のPlanner(planner)からタスク(tasks)を取得
        //  URLパラメータ
        //    なし
        //    通常、Microsoft Graphの各APIではODataクエリによるフィルタとソートが可能だが、
        //    現在Planner関連のAPIではこれがサポートされていない。
        //    例えば $filter=percentComplete lt 100 とすれば未完了データが取得されることが期待されるが、
        //    現段階ではフィルタが行われず全データが取得されてしまう。
        //    よって表示するデータのフィルタ・ソート処理はクライアントサイドで行う必要がある。
        //            
        //  リクエストヘッダ
        //    なし
        return client
          .api("me/planner/tasks/")
          .get((error: GraphError, datas: any, rawResponse?: any) => {
            callBack(error, datas);
          });
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
