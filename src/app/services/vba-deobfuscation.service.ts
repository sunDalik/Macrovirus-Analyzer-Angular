import {Injectable} from '@angular/core';
import {VbaAnalysisService} from "./vba-analysis.service";

@Injectable({
  providedIn: 'root'
})
export class VbaDeobfuscationService {

  constructor(private analysisService: VbaAnalysisService) {
  }


  readonly blocks = [{start: "Sub", end: "End"},
    {start: "Function", end: "End"},
    {start: "Do", end: "Loop"},
    {start: "For", end: "Next"},
    {start: "If", end: "End"},
    {start: "Select", end: "End"},
    {start: "SyncLock", end: "End"},
    {start: "Try", end: "End"},
    {start: "While", end: "End"},
    {start: "With", end: "End"}
  ];

  deobfuscateCode(code: string,
                  doRemoveComments: boolean,
                  doRenameVariables: boolean,
                  doDeobfuscateString: boolean,) {
    let deobfuscatedCode = code;
    //TODO allow preserving original indentation?
    deobfuscatedCode = this.shrinkSpaces(deobfuscatedCode);
    if (doRemoveComments) deobfuscatedCode = this.analysisService.removeComments(deobfuscatedCode);
    deobfuscatedCode = this.analysisService.removeColonDelimiters(deobfuscatedCode);
    if (doRenameVariables) deobfuscatedCode = this.renameVariables(deobfuscatedCode);
    if (doDeobfuscateString) deobfuscatedCode = this.deobfuscateStrings(deobfuscatedCode);
    deobfuscatedCode = this.indentCode(deobfuscatedCode);
    return deobfuscatedCode;
  }

  shrinkSpaces(code: string) {
    const codeLines = code.split("\n");
    for (let i = codeLines.length - 1; i >= 0; i--) {
      codeLines[i] = codeLines[i].trim();
      if (codeLines[i] === "") codeLines.splice(i, 1);
    }
    code = codeLines.join("\n");

    //code = code.replace(/\t/g, ' ');
    //code = code.replace(/  +/g, ' ');
    return code;
  }

  replaceKeyword(text: string, keyword: string, new_keyword: string, i = false) {
    return text.replace(this.analysisService.keywordRegex(keyword, i), "$<before>" + new_keyword);
  }

  indentCode(code: string) {
    //TODO allow changing indent symbol (e.g. 2 spaces instead of 4)
    const blockIndent = "    ";
    let blockStack = [];

    const codeLines = code.split("\n");
    for (let i = 0; i < codeLines.length; i++) {
      let lvlMod = 0;
      codeLines[i] = codeLines[i].trim();
      const matchResult = codeLines[i].match(new RegExp("^[ \\t]*((Private|Public)[ \\t]+)?(?<firstWord>\\b[A-Za-z]+?\\b).*?$"));
      if (matchResult) {
        const firstWord = matchResult.groups?.['firstWord'];
        const blockStart = this.blocks.find(b => b.start === firstWord);
        if (blockStart) {
          blockStack.push(blockStart);
          lvlMod = -1;
        } else if (blockStack.length > 0 && blockStack[blockStack.length - 1].end === firstWord) {
          const poppedBlock = blockStack.pop();
          if (poppedBlock?.start === "Sub" || poppedBlock?.start === "Function") {
            codeLines.splice(i + 1, 0, "");
          }
        }
      }
      for (let lvl = 0; lvl < blockStack.length + lvlMod; lvl++) {
        codeLines[i] = blockIndent + codeLines[i];
      }
    }
    code = codeLines.join("\n");
    return code;
  }

  renameVariables(code: string) {
    //todo block-aware renaming?
    //todo type-aware names?
    //todo function arguments renaming
    const functions = [];
    const newFuncs = [];

    let codeLines = code.split("\n");
    for (const line of codeLines) {
      const matchResult = line.match(this.analysisService.functionDeclarationRegex);
      if (matchResult) {
        functions.push(matchResult.groups?.['functionName']);
      }
    }

    let iterator = 0;
    for (const func of functions) {
      if (func === undefined) continue;
      if (this.analysisService.isAutoExec(func)) continue;
      const new_func_name = "fun_" + iterator;
      for (let i = 0; i < codeLines.length; i++) {
        codeLines[i] = this.replaceKeyword(codeLines[i], func, new_func_name, true);
      }
      newFuncs.push(new_func_name);
      iterator++;
    }

    const variables = [];

    for (const line of codeLines) {
      const matchResult = line.match(this.analysisService.variableDeclarationRegex);
      if (matchResult) {
        if (matchResult.groups?.['variableName']) {
          variables.push(matchResult.groups?.['variableName']);
        } else if (matchResult.groups?.['variableName2']) {
          variables.push(matchResult.groups?.['variableName2']);
        }
      }
    }

    iterator = 0;
    for (const variable of variables) {
      if (newFuncs.includes(variable)) continue;
      const new_var_name = "var_" + iterator;
      for (let i = 0; i < codeLines.length; i++) {
        codeLines[i] = this.replaceKeyword(codeLines[i], variable, new_var_name, true);
      }
      iterator++;
    }

    return codeLines.join("\n");
  }

// TODO very very bad. Can find functions inside strings and doesnt support nested functions
  deobfuscateStrings(code: string) {
    let codeLines = code.split("\n");
    for (let i = 0; i < codeLines.length; i++) {
      codeLines[i] = codeLines[i].replace(this.analysisService.funcCallRegex("Chr"), (match, p1, offset, string, groups) => {
        const arg = Number(groups.args);
        if (Number.isNaN(arg)) return match;
        else return "\"" + this.Chr(arg).toString() + "\"";
      });
      codeLines[i] = codeLines[i].replace(this.analysisService.funcCallRegex("ChrW"), (match, p1, offset, string, groups) => {
        const arg = Number(groups.args);
        if (Number.isNaN(arg)) return match;
        else return "\"" + this.Chr(arg).toString() + "\"";
      });
      codeLines[i] = codeLines[i].replace(this.analysisService.funcCallRegex("Asc"), (match, p1, offset, string, groups) => {
        const arg = groups.args;
        if (arg[0] === "\"" && arg[arg.length - 1] === "\"") return this.Asc(arg).toString();
        else return match;
      });
    }
    return codeLines.join("\n");
  }

  Chr(code: number) {
    return String.fromCharCode(code);
  }

  Asc(char: string) {
    return char.charCodeAt(0);
  }
}
