export interface Welcome {
    collection: Collection[];
    pagination: Pagination;
    self:       string;
}

export interface Collection {
    accountNumber:      number;
    accountType:        AccountType;
    balance:            number;
    blockDirectEntries: boolean;
    debitCredit:        DebitCredit;
    name:               string;
    accountingYears:    string;
    self:               string;
    vatAccount?:        VatAccount;
    totalFromAccount?:  Account;
    barred?:            boolean;
    contraAccount?:     Account;
    accountsSummed?:    AccountsSummed[];
    openingAccount?:    Account;
}


export enum AccountType {
  Heading = "heading",
  HeadingStart = "headingStart",
  ProfitAndLoss = "profitAndLoss",
  Status = "status",
  SumInterval = "sumInterval",
  TotalFrom = "totalFrom",
  find = "find"
}

export interface AccountsSummed {
    fromAccount: Account;
    toAccount:   Account;
}

export interface Account {
    accountNumber: number;
    self:          string;
}

export enum DebitCredit {
    Credit = "credit",
    Debit = "debit",
}

export interface VatAccount {
    vatCode: VatCode;
    self:    string;
}

export enum VatCode {
    I25 = "I25",
    I50 = "I50",
    Ieuy = "IEUY",
    Ivy = "IVY",
    Rep = "REP",
    U25 = "U25",
    Ueuy = "UEUY",
    Uvy = "UVY",
}

export interface Pagination {
    skipPages:            number;
    pageSize:             number;
    maxPageSizeAllowed:   number;
    results:              number;
    resultsWithoutFilter: number;
    firstPage:            string;
    lastPage:             string;
}
