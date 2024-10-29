import { sheetExistChecker } from "./sheetExistChecker";

// GASから参照したい変数はglobalオブジェクトに渡してあげる必要がある
(global as any).sheetExistChecker = sheetExistChecker;
