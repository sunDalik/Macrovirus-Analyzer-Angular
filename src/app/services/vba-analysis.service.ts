import {Injectable} from '@angular/core';
import {OleFile, VBAFunction} from "../models/ole-file";

export enum FuncType {
  Normal,
  DLL
}

@Injectable({
  providedIn: 'root'
})
export class VbaAnalysisService {

  constructor() {
  }

  readonly varName = "[A-Za-z][A-Za-z0-9_\-]*";

  readonly functionDeclarationRegex = new RegExp(`^[ \\t]*((Public|Private)[ \\t]+)?(Sub|Function)[ \\t]+(?<functionName>${this.varName})[ \\t]*\\(.*\\)`);
  readonly functionEndRegex = new RegExp(`^[ \\t]*End[ \\t]*(Sub|Function)[ \\t]*$`, 'm');
  keywordRegex = (keyword: string, i = false) => {
    let flags = "g";
    if (i) flags += "i";
    return new RegExp(`^(?<before>(?:[^"]*?"[^"]*?")*?[^"]*?)\\b${keyword}\\b`, flags);
  };

  funcCallRegex = (funcName: string) => {
    return new RegExp(`\\b${funcName}[ \\t]*\\((?<args>.*?)\\)`, "gi");
  };

  stringRegex = (str: string, i = false) => {
    let flags = "g";
    if (i) flags += "i";
    return new RegExp(`\"${str}\"`, flags);
  };

  readonly dllDeclarationRegex = new RegExp(`^[ \\t]*((Public|Private)[ \\t]+)?Declare[ \\t]+(Sub|Function)[ \\t]+(?<functionName>${this.varName})`);

  readonly variableDeclarationRegex = new RegExp(`(^[ \\t]*(Set|Dim)[ \\t]+(?<variableName>${this.varName}).*?$)|(^[ \\t]*(?<variableName2>${this.varName})([ \\t]*\\(.+?\\))?[ \\t]*=.*?$)`);
  readonly fullLineCommentRegex = new RegExp(`^[ \\t]*'.*`);
  readonly commentRegex = new RegExp(`^(?<before_comment>(?:[^"]*?"[^"]*?")*?[^"]*?)(?<comment>'.*)`);
  readonly lineBreakRegex = new RegExp(` _(\r\n|\r|\n)`, "g");
  readonly colonDelimiterRegex = new RegExp(`:(?!=)`, "g");

  readonly autoExecFunctions = [
    /Workbook_Open/i,
    /Document_Open/i,
    /Document_Close/i,
    /Auto_?Open/i,
    /Workbook_BeforeClose/i,
    /Main/i, //TODO check if code module name contains autoexec AND has Main sub

    // Some ActiveX stuff that I don't understand
    /\w+_Layout/i,
    /\w+_Resize/i,
    /\w+_GotFocus/i,
    /\w+_LostFocus/i,
    /\w+_MouseHover/i,
    /\w+_Click/i,
    /\w+_Change/i,
    /\w+_BeforeNavigate2/i,
    /\w+_BeforeScriptExecute/i,
    /\w+_DocumentComplete/i,
    /\w+_DownloadBegin/i,
    /\w+_DownloadComplete/i,
    /\w+_FileDownload/i,
    /\w+_NavigateComplete2/i,
    /\w+_NavigateError/i,
    /\w+_ProgressChange/i,
    /\w+_PropertyChange/i,
    /\w+_SetSecureLockIcon/i,
    /\w+_StatusTextChange/i,
    /\w+_TitleChange/i,
    /\w+_MouseMove/i,
    /\w+_MouseEnter/i,
    /\w+_MouseLeave/i,
    /\w+_OnConnecting/i,
    /\w+_FollowHyperlink/i,
    /\w+_ContentControlOnEnter/i
  ];

  readonly suspiciousWords = [
    {word: "Kill", description: "Deletes files"},
    {word: "CreateTextFile", description: "Writes to file system"},
    {word: "Open", description: "Opens a file"},
    {word: "Get", description: "Can create shell to execute external programs"},
    {word: "Shell", description: "Runs an executable or a system command"},
    {word: "Create", description: "Can create shell to execute external programs"}, // https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/create-method-in-class-win32-process
    {word: "CreateObject", description: "Can create shell to execute external programs"},
    {word: "GetObject", description: "Can get shell to execute external programs"}
  ];

  readonly suspiciousStrings = [
    {string: "Microsoft.XMLHTTP", description: "Can download a file from the internet"},
    {string: "Wscript.Shell", description: "Runs an executable or a system command"}
  ];

  readonly suspiciousRegex = [
    {regex: /Options\..+[ \t]*=/, word: "Options", description: "Modifies word settings"}
  ];

  analyzeFile(oleFile: OleFile): void {
    oleFile.suspiciousKeywords = [];
    oleFile.stompedKeywords = [];

    if (oleFile.readError) return;

    if (oleFile.macroModules.length === 0) return;

    oleFile.VBAFunctions = this.parseVBAFunctions(oleFile);

    let safe = true;

    const stompingDetectionResult = this.detectVBAStomping(oleFile);
    for (const res of stompingDetectionResult) {
      safe = false;
      oleFile.stompedKeywords.push(res);
    }


    for (const func of oleFile.VBAFunctions) {
      if (!this.isAutoExec(func.name)) continue;
      let fullBody = func.body.join("\n") + "\n";
      const foundWords = [];
      const funcDependencies = this.getAllDependencies(func, oleFile.VBAFunctions);
      for (const dependency of funcDependencies) {
        fullBody += dependency.body.join("\n") + "\n";
      }
      fullBody = this.prepareForAnalysis(fullBody);

      if (funcDependencies.some(f => f.type === FuncType.DLL)) {
        foundWords.push({word: "Lib", description: "Executes DLLs"});
        safe = false;
      }

      for (const word of this.suspiciousWords) {
        if (this.keywordRegex(word.word).test(fullBody)) {
          foundWords.push(word);
          safe = false;
        }
      }

      for (const str of this.suspiciousStrings) {
        if (this.stringRegex(str.string, true).test(fullBody)) {
          foundWords.push({word: "\"" + str.string + "\"", description: str.description});
          safe = false;
        }
      }

      for (const susReg of this.suspiciousRegex) {
        if (susReg.regex.test(fullBody)) {
          foundWords.push(susReg);
          safe = false;
        }
      }

      if (foundWords.length > 0) oleFile.suspiciousKeywords.push({func: func, words: foundWords})
    }

    oleFile.isMalicious = !safe;
  }

  parseVBAFunctions(oleFile: OleFile) {
    const VBAFunctions: Array<VBAFunction> = [];
    let id = 0;
    for (const module of oleFile.macroModules) {
      const moduleCode = this.prepareForAnalysis(module.sourceCode);
      let currentFunction: VBAFunction | null = null;
      for (const line of moduleCode.split("\n")) {
        if (currentFunction !== null) {
          if (this.functionEndRegex.test(line)) {
            currentFunction = null;
          } else {
            currentFunction.body.push(line);
          }
        } else {
          let matchResult = line.match(this.functionDeclarationRegex);
          if (matchResult && matchResult.groups) {
            currentFunction = {
              id: id++,
              name: matchResult.groups?.['functionName'],
              dependencies: [],
              body: [],
              type: FuncType.Normal
            };
            VBAFunctions.push(currentFunction);
            continue;
          }

          matchResult = line.match(this.dllDeclarationRegex);
          if (matchResult && matchResult.groups) {
            VBAFunctions.push({
              id: id++,
              name: matchResult.groups?.['functionName'],
              dependencies: [],
              body: [],
              type: FuncType.DLL
            });
          }
        }
      }
    }

    for (const func of VBAFunctions) {
      const body = func.body.join("\n");
      for (const func2 of VBAFunctions) {
        if (func === func2) continue;
        if (this.keywordRegex(func2.name).test(body)) {
          func.dependencies.push(func2.id);
        }
      }
    }

    //console.log(VBAFunctions);
    return VBAFunctions;
  }

  prepareForAnalysis(code: string) {
    code = this.collapseLongLines(code);
    code = this.removeComments(code);
    code = this.removeColonDelimiters(code);
    return code;
  }

  isAutoExec(func: string) {
    return this.autoExecFunctions.some(f => f.test(func));
  }

  getAllDependencies(func: VBAFunction, VBAFunctions: Array<VBAFunction>) {
    const findDependencies = (currFunc: VBAFunction, fullDependencies: Array<VBAFunction>) => {
      for (const id of currFunc.dependencies) {
        const dep = VBAFunctions.find(f => f.id === id);
        if (dep && !fullDependencies.includes(dep)) {
          fullDependencies.push(dep);
          findDependencies(dep, fullDependencies);
        }
      }
      return fullDependencies;
    };

    return findDependencies(func, []);
  }

  detectVBAStomping(oleFile: OleFile) {
    const result = [];
    let fullSourceCode = "";
    for (const module of oleFile.macroModules) {
      fullSourceCode += module.sourceCode + " ";
    }
    for (const module of oleFile.macroModules) {
      const keywords = new Set<string>();
      for (const command of module.pcode) {
        for (const line of command.split("\n")) {
          const i = line.indexOf(" ");
          const tokens = [line.slice(0, i), line.slice(i + 1)];
          const mnemonic = tokens[0];

          let args = '';
          if (tokens.length === 2) {
            args = tokens[1];
          }

          if (['ArgsCall', 'ArgsLd', 'St', 'Ld', 'MemSt', 'Label'].includes(mnemonic)) {
            if (args.startsWith("(Call) ")) {
              args = args.slice(7);
            }
            const keyword = args.split(" ")[0];
            if (!keyword.startsWith("id_")) {
              keywords.add(keyword);
            }
          }

          if (mnemonic === 'LitStr') {
            const i = line.indexOf(" ");
            let str = args.slice(i + 1);
            if (str.length >= 2) {
              str = str.slice(1, str.length - 1);
              str = str.replaceAll('"', '""');
              str = '"' + str + '"';
            }
            keywords.add(str);
          }
        }
      }

      for (const keyword of keywords.values()) {
        if (keyword === "undefined") continue;
        if (!fullSourceCode.toLowerCase().includes(keyword.toLowerCase())) {
          result.push({module: module, keyword: keyword});
        }
      }
    }
    return result;
  }


  removeComments(code: string) {
    let newCodeLines = [];
    for (const line of code.split("\n")) {
      if (!this.fullLineCommentRegex.test(line)) {
        newCodeLines.push(line.replace(this.commentRegex, "$<before_comment>"));
      }
    }
    return newCodeLines.join("\n");
  }

  removeColonDelimiters(code: string) {
    code = code.replace(this.colonDelimiterRegex, "\n");
    return code;
  }

  collapseLongLines(code: string) {
    return code.replaceAll(this.lineBreakRegex, " ");
  }
}
