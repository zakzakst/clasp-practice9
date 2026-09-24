export type FormItem = {
  page: string;
  label: string;
  selector: string;
  type: "text" | "checkbox";
  defaultValue: string | boolean;
};

export type SheetInfos = {
  sheetName: string;
  formItems: FormItem[];
}[];
