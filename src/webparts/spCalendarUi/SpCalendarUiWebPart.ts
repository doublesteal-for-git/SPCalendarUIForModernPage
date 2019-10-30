import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  SPHttpClient,
  SPHttpClientResponse,   
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import styles from './SpCalendarUiWebPart.module.scss';
import * as strings from 'SpCalendarUiWebPartStrings';
//ログクラス
import Logger from './Logger'
//日付フォーマット用
import * as Moment from 'moment';
import { SPList } from '@microsoft/sp-page-context';
import { any, number } from 'prop-types';
import { DOMElement } from 'react';


//インターフェース定義
export interface ISpCalendarUiWebPartProps {
  description: string;
  targetList: string;
};
export interface ISPLists {
  value: ISPList[];
};
export interface ISPList {
  Title: string;
  Id: string;
  Description: string;
};
export interface ISPItems{
  value: ISPItem[];
};
export interface ISPItem{
  Title: string;
  Location: string;
  EventDate: Date;
  EndDate: Date;
  fAllDayEvent: boolean;
  fRecurrence: boolean;
  Duration: number;
  RecurrenceData: string;
};
// 繰り返し用データ
export interface IRecurrenceData{
  repeatInstances: number;
  windowEnd: Date;
  repeatOption: repeatOptions;
  frequencyOption: frequencyOptions;
  frequencyNum: number;
  startMonth: number;
  startDay: number;
  dayOfTheWeek: {[key: number]:boolean};

};
export enum repeatOptions {
  Forever,
  Instance,//回数指定
  WindowEnd//終了日指定
};
export enum frequencyOptions {
  Yearly,
  Monthly,
  Weekly,
  Daily,
  Weekday//平日全て
};
export enum dayOfTheWeekOptions {
  Su,
  Mo,
  Tu,
  We,
  Th,
  Fr,
  Sa
};


export default class SpCalendarUiWebPart extends BaseClientSideWebPart<ISpCalendarUiWebPartProps> {
  private ddLists: IPropertyPaneDropdownOption[] = [{key:'_blank',text:'------Select List------'}];
  private ddListsDropdownDisabled: boolean = true;
  private log = new Logger(Logger.logLevel.NoLogging);
  // 現在の選択年月日（YYYY-MM）
  private selectedDate: string = "";
  // 選択年月の月初と月末
  private startDate: Date;
  private endDate: Date;
  

  public render(): void {
    this.log.debug("対象リストの選択状況："+ this.properties.targetList);
    
    if(this.properties.targetList === "_blank" || this.properties.targetList === ''){
      this.domElement.innerHTML = `
      <div>
        <div class="${ styles.container }">
          <span class="${ styles.title }">Please edit WebParts and select target calendar list.</span>
        </div>
      </div>`;

    } else {
        this._renderCalendar();
    }
    
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
                PropertyPaneDropdown('targetList', {
                  label: strings.TargetListFieldLabel,
                  options: this.ddLists,
                  disabled: this.ddListsDropdownDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
  // #region Webパーツオプション設定処理==============================================
  // サイト内のカレンダーテンプレートで作成されたリストを取得
  private _getCalendarLists(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=BaseTemplate eq 106`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  // カレンダーリスト情報からオプションを更新する
  private _updateDropDownOptions(): void {
    this._getCalendarLists().then((response)=>{
      this.log.debug("リスト取得成功："+ JSON.stringify(response.value));
      this._setDropdownOptions(response.value);  
    }).catch((e:any)=>{
      alert("Unexpected Error : get failed calendar List.\n"+e.message);
      return null;
    });
  }

  private _setDropdownOptions(lists:ISPList[]): void{
    lists.forEach((list:ISPList)=>{
      this.log.debug("リスト情報：" + list.Title);
      this.ddLists.push({key:list.Title,text:list.Title});  
    }); 
    this.ddListsDropdownDisabled = false;
    this.context.propertyPane.refresh();
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.render();
  }
  
  // プロパティ初期設定の既定メソッド
  protected onPropertyPaneConfigurationStart(): void {
    this.log.debug("プロパティコンフィグ構成開始");
    this.log.debug(`Locale情報:${JSON.stringify(this.context.pageContext.cultureInfo)}`);
    this.ddListsDropdownDisabled = !this.ddLists;
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'While getting calendar list');
    // ドロップダウン更新
    this._updateDropDownOptions();
  }
  // #endregion Webパーツオプション設定処理==============================================

  //#region  カレンダーアイテム取得とレンダリング================================================
  // カレンダーリストアイテム取得
  private _getCalendarListItems(baseQuery: any): Promise<ISPItems> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.targetList}')/Items?`+baseQuery, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  // カレンダーリストレンダリング
  private _renderCalendar() : void {
    // 対象日付範囲取得
    if(this.selectedDate==""){
      this.selectedDate = `${new Date().getFullYear()}-${new Date().getMonth()+1}`;
    }
    
    if(this.context.pageContext.cultureInfo.currentCultureName == "ja-JP"){
      Moment.locale("ja");
    }else{
      Moment.locale("en");
    }
    
    this.startDate = this._getMonthStartDate();
    this.endDate = this._getMonthEndDate();
    // 先に対象月のカレンダー生成
    this._preRenderMonthCalendarElement();

    let queryStartDate = this.startDate.toISOString();
    let queryEndDate = this.endDate.toISOString();

    // 対象月に跨るイベント取得クエリ
    var baseQuery: any = `$select=Title,Location,EventDate,EndDate,fAllDayEvent,fRecurrence,Duration,RecurrenceData&$filter=((EventDate ge datetime'${queryStartDate}') and (EndDate le datetime'${queryEndDate}')) or((EndDate ge datetime'${queryEndDate}') and (EventDate le datetime'${queryEndDate}')) or((EndDate ge datetime'${queryStartDate}') and (EventDate le datetime'${queryStartDate}'))&$orderby=EventDate asc`;

    this.log.debug("クエリ："+baseQuery);
    this._getCalendarListItems(baseQuery)
    .then((response)=>{
      response.value.forEach((item: ISPItem)=> {
        // カレンダーにイベント追加
        this._buildEventElement(item);
      });
      
    }).catch((e:any)=>{
      alert("Unexpected Error : get failed calendar List item.\n" + e.message);
      return null;
    });
    
  }
  
  // カレンダー表示月変更
  private changeSelectedDate(option:string){
    let curYear:number = parseInt(this.selectedDate.substring(0,4));
    let curMonth:number = parseInt(this.selectedDate.substring(5,7));
    let changeYear:string ="";
    let changeMonth:string ="";
    if(option=="next"){
      changeMonth =  curMonth == 12 ? "1":(curMonth+1).toString();
      changeYear =   curMonth == 12 ? (curYear+1).toString():curYear.toString();
    }else if(option=="prev"){
      changeMonth =  curMonth == 1 ? "12":(curMonth-1).toString();
      changeYear =   curMonth == 1 ? (curYear-1).toString():curYear.toString();
    }
    changeMonth = changeMonth.length == 1? '0'+changeMonth : changeMonth;
    this.selectedDate = `${changeYear}-${changeMonth}`;
    this._renderCalendar();
  }

  // 月表示のカレンダーDOM生成
  private _preRenderMonthCalendarElement(){
    let calendarDomElem:string = "";
    calendarDomElem = `
    <div class="${styles.spCalendarUi}">
    <div class="${styles.container}">
    <div class="${styles.tableStyle}">
    
    <h2><button class="${styles.button}" id="prevMonthBtn"> Prev Month << </button>
      ${this.selectedDate}  
      <button class="${styles.button}" id="nextMonthBtn"> >> Next Month </button></h2>
    
    <table>
      <tbody>
        <th><nobr><span>${ strings.Sunday }</span></nobr></th>
        <th><nobr><span>${ strings.Monday }</span></nobr></th>
        <th><nobr><span>${ strings.Tuesday }</span></nobr></th>
        <th><nobr><span>${ strings.Wednesday }</span></nobr></th>
        <th><nobr><span>${ strings.Thursday }</span></nobr></th>
        <th><nobr><span>${ strings.Friday }</span></nobr></th>
        <th><nobr><span>${ strings.Saturday }</span></nobr></th>
        <tr>
      `;
      let startDayOfWeek: number = this.startDate.getDay();
      let lastDay: number = this.endDate.getDate();
      let lastDayOfWeek: number = this.endDate.getDay();

      for ( let curDay: number = -startDayOfWeek + 1; curDay <= lastDay; curDay++ ) {
        if(curDay < 1){
          calendarDomElem += "<td></td>";
        }else{
          let id: string = "";
          //各日付列に設定するID（形式：{リスト名}_{選択年月}_{日付}）
          if(curDay < 10){
            id = `${this.properties.targetList}_${this.selectedDate}-0${curDay.toString()}`;
          }else{
            id = `${this.properties.targetList}_${this.selectedDate}-${curDay.toString()}`;
          }
          calendarDomElem += `<td><div>${curDay.toString()}</div><div id="${id}"></div></td>`;
        }
        // 週最終日
        if(( startDayOfWeek + curDay ) % 7 == 0 ) {
          calendarDomElem += `</tr><tr>`;
        }
      }
      // テーブル調整
      for(let i = 1; i < (7 - lastDayOfWeek); i++){
        calendarDomElem += "<td></td>";
      }

      calendarDomElem += "</tbody></table></div></div></div>";
      this.domElement.innerHTML = calendarDomElem;
      // ボタンにクリックイベント付与
      this.domElement.querySelector('#prevMonthBtn').addEventListener('click', () => {  
        this.changeSelectedDate("prev");
      });
      this.domElement.querySelector('#nextMonthBtn').addEventListener('click', () => {  
        this.changeSelectedDate("next");
      }); 

  }

  // イベント生成
  private _buildEventElement(item: ISPItem){
    this.log.debug(JSON.stringify(item));
    // 定期イベントのデータ解析
    let rdata:IRecurrenceData = null;
    if(item.fRecurrence){
      // 第何週目指定には現時点では未対応
      if(item.RecurrenceData.indexOf('ByDay')!=-1){
        return;
      }
      rdata = this._getRecurrenceData(item.RecurrenceData);
      this.log.debug(JSON.stringify(rdata));
    }
    // 表示用にフォーマット
    // 終日予定の場合ローカルタイム、それ以外はUTCがJSONで返ってくる(SPO仕様))
    if(item.fAllDayEvent){
      var eventYear = item.EventDate.toString().substring(0,4);
      var eventMonth = item.EventDate.toString().substring(5,7);
      var eventDay = item.EventDate.toString().substring(8,10);
      var eventTime = item.EventDate.toString().substring(11,16);

      var endYear = item.EndDate.toString().substring(0,4);
      var endMonth = item.EndDate.toString().substring(5,7);
      var endDay = item.EndDate.toString().substring(8,10);
      var endTime = item.EndDate.toString().substring(11,16);
    }else{
      var eventYear = Moment(item.EventDate).format("YYYY");
      var eventMonth = Moment(item.EventDate).format("MM");
      var eventDay = Moment(item.EventDate).format("DD");
      var eventTime = Moment(item.EventDate).format("HH:mm");

      var endYear = Moment(item.EndDate).format("YYYY");
      var endMonth = Moment(item.EndDate).format("MM");
      var endDay = Moment(item.EndDate).format("DD");
      var endTime = Moment(item.EndDate).format("HH:mm");
    }
    
    let diff:number = 0;
    // 差分抽出用
    let sdate: Date = new Date(item.EventDate.toString());
    let edate: Date = new Date(item.EndDate.toString());;
    
    // 日付が跨る場合は差分を計算
    if(`${eventYear}-${eventMonth}-${eventDay}` != `${endYear}-${endMonth}-${endDay}`){
      diff = Math.floor((edate.getTime()-sdate.getTime())/(1000 * 60 * 60 *24));
    }

    let loopStart:number = 1;
    // 表示している対象月より前に開始されるイベントの開始日特定
    if(edate.getTime()-this.startDate.getTime() >= 0){
      // 対象月内に開始されるイベントはEventDateが開始日
      if(eventMonth == endMonth || (this.startDate <= sdate && this.endDate >= sdate)){       
        loopStart = parseInt(eventDay,10);
      }
    }

    //繰り返しイベントの場合は開始日を繰り返しスキーマから特定
    if(item.fRecurrence){
      loopStart = this._getRecStartDay(item,rdata,eventYear,eventMonth,eventDay,diff);
      if(loopStart == null){
        return;
      }
      if(rdata.frequencyOption == frequencyOptions.Yearly || rdata.frequencyOption == frequencyOptions.Monthly){
        diff = 0;
      }
    }
    this.log.debug("loopStart:"+loopStart);
    this.log.debug("diff："+diff.toString());
    //複数の要素にイベントを追加
    let loopCount:number = 0;
    for(let i = loopStart; i <= diff+loopStart; i++){
      // 上書き対象のカレンダー要素IDNo
      let keyDay:number = loopStart + loopCount;
      let keyDayString: string = keyDay < 10 ? `0${keyDay.toString()}` : keyDay.toString();
      
      // イベント終了日を超える場合は終了
      if(parseInt(endMonth,10) == parseInt(this.selectedDate.substring(5,7),10) && keyDay > parseInt(endDay,10)){
        break;
      }

      // 次に進める日付計算
      let curDate = new Date(`${this.selectedDate}-${keyDayString}`);
      let plusCount:number = 0;
      if(item.fRecurrence &&　rdata.frequencyOption == frequencyOptions.Daily){
        plusCount = rdata.frequencyNum;// 繰り返し間隔分次に日付を進める
      // 週最終日なら繰り返し間隔*7+1日進める
      }else if(item.fRecurrence &&　rdata.frequencyOption == frequencyOptions.Weekly && curDate.getDay()==6){
        plusCount = rdata.frequencyNum == 1? 1: (rdata.frequencyNum-1)*7+1;
      }else{
        plusCount = 1;
      }

      // 繰り返し曜日じゃない場合次へ進む  
      if(item.fRecurrence && 
        (rdata.frequencyOption == frequencyOptions.Weekday || 
           rdata.frequencyOption == frequencyOptions.Weekly)){
             
             if(!rdata.dayOfTheWeek[curDate.getDay()]){
              loopCount += plusCount;
              continue;
             }
      }
      // 対象月内のみイベント作成するため31以上はループ不要
      if(keyDay > 31){
        break;
      }
        
      let targetId: string = `#${this.properties.targetList}_${this.selectedDate}-${keyDayString}`;
      // 要素取得
      let targetElem: Element = this.domElement.querySelector(`${targetId}`);
      this.log.debug("key:"+targetId);
      
      if(targetElem == null){
        // オプションごとに次の日付指定
        loopCount += plusCount;
        continue;
      }

      let eventElem: string;
      let insertPosition: InsertPosition;
      let titleMaxLen: number = 10;//イベントタイトル長さMAX値

      // 繰り返しイベントは終了年月は表示しない
      if(item.fRecurrence){
        eventElem= `<span title="${eventTime} - ${endTime}
        Location: ${item.Location}
        Title: ${item.Title}
        ">${item.Title.length >= titleMaxLen ? "&#x21BB; "+item.Title.substring(0,titleMaxLen)+"..." : "&#x21BB; "+item.Title}</span><br>`;
        insertPosition = 'afterbegin';
      // 終日イベントか複数日にまたがるイベントは時間表示なし
      }else if(diff != 0 || item.fAllDayEvent){
        eventElem= `<span title="${eventYear}/${eventMonth}/${eventDay} ${eventTime} - ${endYear}/${endMonth}/${endDay} ${endTime}
        Location: ${item.Location}
        Title: ${item.Title}
        ">${item.Title.length >= titleMaxLen ? item.Title.substring(0,titleMaxLen)+"..." : item.Title}</span><br>`;
        insertPosition = 'afterbegin';
      }else{
        eventElem= `<span title="${eventYear}/${eventMonth}/${eventDay} ${eventTime} - ${endYear}/${endMonth}/${endDay} ${endTime}
        Location: ${item.Location}
        Title: ${item.Title}
        ">${eventTime} ${item.Title.length >= titleMaxLen ? item.Title.substring(0,titleMaxLen)+"..." : item.Title}</span><br>`;
        insertPosition = 'beforeend';
      }
      // 要素に追加
      targetElem.insertAdjacentHTML(insertPosition,eventElem);
      // オプションごとに次の日付指定
      loopCount += plusCount;
    }
  }
  //#endregion カレンダーアイテム取得およびレンダリング================================================

  //#region Util
  // 月初時刻取得：Date型
  private _getMonthStartDate(): Date {
    let curdate = new Date(this.selectedDate);
    return (new Date(curdate.getFullYear(), curdate.getMonth(), 1));
  }

  // 月末時刻取得：Date型
  private _getMonthEndDate(): Date {
    let curdate = new Date(this.selectedDate);
    let lastdate = new Date(curdate.getFullYear(), curdate.getMonth() + 1, 0);
    lastdate.setHours(lastdate.getHours() + 23);
    lastdate.setMinutes(lastdate.getMinutes() + 59);
    return lastdate;
  }

  //繰り返しイベントのスキーマ分解
  private _getRecurrenceData(rdata: string):IRecurrenceData{
    let rData: IRecurrenceData;
    let rpOpt: repeatOptions;
    let rpIns: number;
    let winEnd: Date;
    let freqOpt: frequencyOptions;
    let freqNum: number;
    let month: number;
    let day: number;
    let dayOfTheWeeks: { [key: number]: boolean; } = {
      [dayOfTheWeekOptions.Su]:false,
      [dayOfTheWeekOptions.Mo]:false,
      [dayOfTheWeekOptions.Tu]:false,
      [dayOfTheWeekOptions.We]:false,
      [dayOfTheWeekOptions.Th]:false,
      [dayOfTheWeekOptions.Fr]:false,
      [dayOfTheWeekOptions.Sa]:false
    };

    // 繰り返し回数オプションの判定
    if(rdata.indexOf("<repeatInstances>")!= -1){
      rpOpt = repeatOptions.Instance;
      //繰り返し回数抽出
      rpIns = parseInt(rdata.substring(rdata.indexOf("<repeatInstances>")+"<repeatInstances>".length,rdata.indexOf("</repeatInstances>")),10);
    }else if(rdata.indexOf("<windowEnd>")!= -1){
      rpOpt = repeatOptions.WindowEnd;//繰り返し最終日指定
      winEnd = new Date(rdata.substring(rdata.indexOf("<windowEnd>")+"<windowEnd>".length,rdata.indexOf("</windowEnd>")));
    }else{
      rpOpt = repeatOptions.Forever;
    }
    
    //繰り返しオプション判定
    let startFreqOptNum: number = rdata.indexOf("<repeat><")+"<repeat><".length;
    let startFreqOpt: string =  rdata.substring(startFreqOptNum,rdata.indexOf(" ",startFreqOptNum));
    
    let startdNum:number = rdata.indexOf('day="')+'day="'.length;
    if(startdNum!=-1){
      day = parseInt(rdata.substring(startdNum,rdata.indexOf('"',startdNum)),10);
    }

    if(startFreqOpt == "yearly"){ 
      freqOpt = frequencyOptions.Yearly;
      let startMoNum:number = rdata.indexOf('month="')+'month="'.length;
      month = parseInt(rdata.substring(startMoNum,rdata.indexOf('"',startMoNum)),10);
    }
    else if(startFreqOpt == "monthly"){
      freqOpt = frequencyOptions.Monthly;
    }
    else if(startFreqOpt == "weekly"){
      freqOpt = frequencyOptions.Weekly;
    }
    else{ 
      if(rdata.indexOf('weekday="TRUE"')!=-1){
        freqOpt = frequencyOptions.Weekday;
        freqNum = 1;
      }else{
        freqOpt = frequencyOptions.Daily;
      }    
    }
    // 繰り返し間隔
    let startFreqNum: number = rdata.indexOf('Frequency="')+'Frequency="'.length;
    if(startFreqNum!=-1){
      freqNum =  parseInt(rdata.substring(startFreqNum,rdata.indexOf('"',startFreqNum)),10);
    }
    
    // 曜日特定
    if(freqOpt != frequencyOptions.Yearly && freqOpt != frequencyOptions.Daily){
      if(rdata.indexOf('su="TRUE"')!=-1){ dayOfTheWeeks[dayOfTheWeekOptions.Su]=true }
      if(rdata.indexOf('mo="TRUE"')!=-1 || freqOpt == frequencyOptions.Weekday){ dayOfTheWeeks[dayOfTheWeekOptions.Mo]=true }
      if(rdata.indexOf('tu="TRUE"')!=-1 || freqOpt == frequencyOptions.Weekday){ dayOfTheWeeks[dayOfTheWeekOptions.Tu]=true }
      if(rdata.indexOf('we="TRUE"')!=-1 || freqOpt == frequencyOptions.Weekday){ dayOfTheWeeks[dayOfTheWeekOptions.We]=true }
      if(rdata.indexOf('th="TRUE"')!=-1 || freqOpt == frequencyOptions.Weekday){ dayOfTheWeeks[dayOfTheWeekOptions.Th]=true }
      if(rdata.indexOf('fr="TRUE"')!=-1 || freqOpt == frequencyOptions.Weekday){ dayOfTheWeeks[dayOfTheWeekOptions.Fr]=true }
      if(rdata.indexOf('sa="TRUE"')!=-1){ dayOfTheWeeks[dayOfTheWeekOptions.Sa]=true }
    }

    // 取得した各プロパティ設定
    rData = {
      repeatOption: rpOpt,
      repeatInstances: rpIns,
      windowEnd: winEnd,
      frequencyOption: freqOpt,
      frequencyNum: freqNum,
      startMonth: month,
      startDay: day,
      dayOfTheWeek: dayOfTheWeeks
    };
    
    return rData;
  }

  // 繰り返しイベント開始日と差分の判定
  private _getRecStartDay(item: ISPItem, rdata: IRecurrenceData, eventYear: string, eventMonth: string, eventDay: string, diff: number): number{
    let loopStart:number = null;
    let isEventEqTarget:boolean = false;
    let maxLoop = 1;// マックスループ

    // イベント開始月が対象月と同じ
    if(parseInt(eventYear,10) == this.startDate.getFullYear() && parseInt(eventMonth,10) == (this.startDate.getMonth() +1)){
      isEventEqTarget = true;
    }
    // 年次イベント
    if(rdata.frequencyOption == frequencyOptions.Yearly){
      // 開始月が対象月ではない場合はレンダリング不要
      if(rdata.startMonth != (this.startDate.getMonth()+1)){
        return null;
      }else{
        loopStart = rdata.startDay;
      }
    // 月次イベント
    }else if(rdata.frequencyOption == frequencyOptions.Monthly){
      // イベント開始月が対象月と同じ
      if(isEventEqTarget){
        loopStart = rdata.startDay;
      }else{
        maxLoop = diff/(rdata.frequencyNum*28);// 28日未満の月はないので28で割れば確実にループできる
        let nextDate = new Date(`${eventYear}-${eventMonth}`);
        for(let i = 1; i <= maxLoop; i++){
          nextDate.setMonth(nextDate.getMonth() + rdata.frequencyNum);// 次のイベントへ進める
          // 対象月以下
          if(this.startDate.getFullYear() >= nextDate.getFullYear() && this.startDate.getMonth()+1 > nextDate.getMonth()+1){
            continue;
          // 対象月 
          }else if(this.startDate.getFullYear() == nextDate.getFullYear() && this.startDate.getMonth()+1 == nextDate.getMonth()+1){
            loopStart = rdata.startDay;
            break;
          // 次のイベントが対象月をオーバーしたのでレンダリング不要
          }else{
            return null;
          }
        }
      }
    // 週イベントまたは日時イベント
    }else if(rdata.frequencyOption == frequencyOptions.Weekly || rdata.frequencyOption == frequencyOptions.Daily|| rdata.frequencyOption == frequencyOptions.Weekday){
      // イベント開始月が対象月と同じ
      if(isEventEqTarget){
        loopStart = parseInt(eventDay);
      }else{
        let nextDate = new Date(`${eventYear}-${eventMonth}-${eventDay}`);
        if(rdata.frequencyOption == frequencyOptions.Weekly){
          maxLoop = diff/(rdata.frequencyNum == 1? 1 : (rdata.frequencyNum-1)*7);
        }else{
          maxLoop = diff/rdata.frequencyNum;
        }
        // 対象の月になるまで日付を進める     
        for(let i = 1; i <= maxLoop; i++){
          // 週の開始日設定
          let curDay:number = nextDate.getDay()
          if(rdata.frequencyOption == frequencyOptions.Weekly){
            let plusCount: number = rdata.frequencyNum == 1? (6-curDay+1): (rdata.frequencyNum-1)*7+(6-curDay+1);
            nextDate.setDate(nextDate.getDate() + plusCount);// 次のイベントへ進める
          }else{
            nextDate.setDate(nextDate.getDate() + rdata.frequencyNum);
          }
          curDay = nextDate.getDay()
          // 対象月以下
          if(this.startDate.getFullYear() >= nextDate.getFullYear() && this.startDate.getMonth()+1 > nextDate.getMonth()+1){
            // 週初めが対象月でなくても、週のいずれかで対象月になる場合を考慮
            let plusDay:number = 6-curDay;
            if(plusDay==0){
              continue;// 週最終日なので次の週へ
            }else if(rdata.frequencyOption == frequencyOptions.Weekly){
              let lastDate = new Date(`${nextDate.getFullYear()}-${nextDate.getMonth()+1}-${nextDate.getDate()}`);
              let breakflg:boolean = false;
              for(let k = 1; k <= plusDay; k++){
                lastDate.setDate(lastDate.getDate() + 1);
                if(this.startDate.getMonth()+1 == lastDate.getMonth()+1){
                  loopStart = lastDate.getDate();
                  breakflg = true;
                  break;
                }
              }
              if(breakflg){
                break;
              }
            }
            continue;
          // 対象月 
          }else if(this.startDate.getFullYear() == nextDate.getFullYear() && this.startDate.getMonth()+1 == nextDate.getMonth()+1){
            
            // 日曜なら開始日として設定
            if(curDay == 0){ 
              loopStart = nextDate.getDate();
              break;
            // 日曜でない場合週の開始日まで日付を戻す
            }else{             
              for(let j = curDay; j > 0; j--){
                let prevDate = new Date(`${nextDate.getFullYear()}-${nextDate.getMonth()+1}-${nextDate.getDate()}`); 
                prevDate.setDate(nextDate.getDate() - j);
                // 日付を戻した際に月を跨ぐ場合はスキップ
                if(this.startDate.getMonth()+1 != prevDate.getMonth()+1){
                  continue;
                }else{
                  loopStart = prevDate.getDate();
                  break;
                }         
              }
            }        
            break;
          // 次のイベントが対象月をオーバーしたのでレンダリング不要
          }else{
            return null;
          }
        }
      }
    }
    return loopStart;
  }
  //#endregion Util
}

