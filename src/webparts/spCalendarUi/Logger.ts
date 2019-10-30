class Logger {
    level: Logger.logLevel;
    //ログフォーマット
    format: string = "YYYY/MM/dd hh:mm:ss [{loglevel}] {msg}";
    constructor(level: Logger.logLevel){
        this.level = level;
    }
    //ログメッセージフォーマット
    formatMessage(message,level): string{
        let date = new Date();
        let res = "";
        res = this.format.replace(/YYYY/, date.getFullYear().toString());
        res = res.replace(/MM/, (date.getMonth()+1).toString());
        res = res.replace(/dd/, date.getDate().toString());
        res = res.replace(/hh/, date.getHours().toString());
        res = res.replace(/mm/, date.getMinutes().toString());
        res = res.replace(/ss/, date.getSeconds().toString());
        res = res.replace(/{loglevel}/,level);
        res = res.replace(/{msg}/,message);
        return res;
    }

    debug(message: string): void{
        if(this.level == Logger.logLevel.Debug){
            console.log(this.formatMessage(message,"debug"));
        }
    }
    info(message: string): void{
        if(this.level >= Logger.logLevel.Info){
            console.log(this.formatMessage(message,"info"));
        }
    }
    warn(message: string): void{
        if(this.level >= Logger.logLevel.Warn){
            console.log(this.formatMessage(message,"warn"));
        }
    }
    error(message: string): void{
        if(this.level >= Logger.logLevel.Error){
            console.log(this.formatMessage(message,"error"));
        }
    }

}

module Logger {
    export enum logLevel
    {
        NoLogging,
        Error,
        Warn,
        Info,
        Debug
    }
}

export default Logger;