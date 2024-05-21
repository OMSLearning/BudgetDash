import * as React from 'react';
import styles from './BudgetDash.module.scss';
import type { IBudgetDashProps } from './IBudgetDashProps';
import {AccountType} from './IBudgetFetch';
import * as FluentUI from '@fluentui/react-components';
import { formatNumber } from './utility/formatNumber';
//import { Label, } from '@fluentui/react';
import { Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Title3, Toolbar, ToolbarButton, } from '@fluentui/react-components';
import {AddSquare16Regular, ArrowNextRegular, ArrowPreviousRegular, Delete16Regular, Edit16Regular, MoreHorizontal24Filled,} from "@fluentui/react-icons";
import { DefaultButton, PrimaryButton, TextField, Dialog, DialogType, DialogFooter, ComboBox, Toggle,  } from '@fluentui/react';
//import type { ToolbarProps } from "@fluentui/react-components";




interface IComponentState {
  accountNumbers: number[];
  accounts: AccountDetail[];
  budgetSum: number;
  currentMonth: string;
  formAccountNumber?: string;
  formYear?: string;
  formMonth?: string;
  formBudget?: string;
  isDialogVisibleBudget: boolean;
  isDialogVisibleEditBudget: boolean;
  isDialogVisibleKonti: boolean;
  isDialogVisibleDeleteKonti: boolean;
  regnSum: string;
  accountName: string; // til Title i budget listen
  accountType: string;
  hoveredRow: number | null;
  loading: boolean,
  kontiOptions: 'yes' | 'standard';
  kontiToDelete: number | null;
  kontiToEditBudget: number | null;
}


interface IComboBoxOption {
  key: string | number;
  text: string;
}

interface IListItemFields {
  Kontonr: number;
  Kontonavn: string;
  BoldStyle: string;
  AddSpacer: string;
}


interface IListItem {
  id: string;
  fields: IListItemFields;
}

interface IListItemsResponse {
  value: IListItem[];
  accDetails: AccountDetail[],
  
}

interface AccountDetail {
  kontiId?: string; 
  accountNumber: number;
  accountType: AccountType;
  balance: number;
  name: string;
  budget: number;
  budgetYear: number;
  budgetMonth: string;
  achieved?: number;
  regnSum: string;
  deviation: number;
  addSpacer: string;
  boldStyle: string; 
  debitCredit: string;
  budgetText: string;
  specificMonthBudget?: number;
  group?: number;
}


export default class BudgetDash extends React.Component<IBudgetDashProps, IComponentState, IListItemsResponse> {
constructor(props: IBudgetDashProps){
  super(props)
  this.state={
    accountNumbers: [],
    accounts: [],
    budgetSum:0,
    currentMonth: this.getMonthName(new Date().getMonth() + 1),
    isDialogVisibleBudget: false,
    isDialogVisibleEditBudget: false,
    isDialogVisibleKonti: false,
    isDialogVisibleDeleteKonti: false,
    hoveredRow: null,
    formAccountNumber: '',
    formYear: '',
    formMonth: '',
    formBudget: '',
    regnSum: "No", 
    accountName: "",
    accountType: "",
    kontiOptions: 'standard',
    kontiToDelete: null,
    kontiToEditBudget: null,
    loading: true,
  };
  this.nextMonthButton = this.nextMonthButton.bind(this);
  this.prevMonthButton = this.prevMonthButton.bind(this);
}
cachedAccountDetails: AccountDetail[] = [];

async componentDidMount() {
  try {
    this.cachedAccountDetails = await this.fetchAccountNumbers();
    const currentDate = new Date();
    const currentMonth = currentDate.getMonth() + 1; // JavaScript/Typescript months are zero-indexed
    const currentYear = currentDate.getFullYear().toString();
    await this.combineAccountInfo(currentMonth, currentYear);
    const accountNumbers = this.cachedAccountDetails.map(account => account.accountNumber);
    if (accountNumbers.length > 0) {
      await this.loadBudgetData(accountNumbers, currentMonth.toString(), currentYear);
    }
  } catch (error) {
    console.error("Failed to initialize account data:", error);
  }
}


processBudgetData(budgetData: any[]) {
  const updatedAccounts = this.state.accounts.map(account => {
    const newBudgetData = budgetData.find((item: { accountNumber: number; }) => item.accountNumber === account.accountNumber);
    return newBudgetData ? { ...account, ...newBudgetData } : { ...account, budget: "No data" };
  });

  this.setState({ accounts: updatedAccounts, loading: false });
}



async loadBudgetData(accountNumbers: number[], month: string, year: string) {
  let dataToUse = await this.fetchBudgetFromSPList(accountNumbers, month, year);
  if (dataToUse.length === 0 || dataToUse.every((item: { budget: number; }) => !item.budget && item.budget !== 0)) {
    dataToUse = await this.fetchBudgetFromSPList(accountNumbers, "*", year);
  }
  this.processBudgetData(dataToUse);
}




async fetchAccountNumbers(): Promise<AccountDetail[]> {
  if (!this.props.selectedKontiList) {
    throw new Error("Selected list ID is not provided");
  }
  const client = await this.props.context.msGraphClientFactory.getClient("3");
  const siteUrl = this.props.context.pageContext.web.serverRelativeUrl;
  const hostname = window.location.hostname;
  const siteId = await client.api(`/sites/${hostname}:${siteUrl}`).get().then(response => response.id);
  const listId = this.props.selectedKontiList;

  const itemsResponse = await client.api(`/sites/${siteId}/lists/${listId}/items`).expand('fields($select=Kontonr,BoldStyle,AddSpacer,RegnSum)').get();
  return itemsResponse.value.map((item: { fields: { Kontonr: number; BoldStyle: string; AddSpacer: string; RegnSum: string; }; }) => ({
    accountNumber: item.fields.Kontonr,
    boldStyle: item.fields.BoldStyle,
    addSpacer: item.fields.AddSpacer,
    regnSum: item.fields.RegnSum,
  }));
}


fetchAccountDetails = async (accountNumbers: number[]) => {
  const filterQuery = `accountNumber$in:[${accountNumbers.join(',')}]`;
  const url = `https://restapi.e-conomic.com/accounts?filter=${encodeURIComponent(filterQuery)}`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      'X-AppSecretToken': "1NglExETODyTzOEr5aqNxmhjmE9VjWOli2lhoEcao5g",
      'X-AgreementGrantToken': "CwRfoodzmqD2b1cRsJpF27zwaCILiSNrjlC7JejFtB81",
      'Content-Type': "application/json"
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch account details for numbers: ${accountNumbers.join(', ')}`);
  }
  
  const data = await response.json();
  console.log("Account data:", data);
  return data.collection.map((account: { accountNumber: any; accountType: any; balance: number; name: any; debitCredit: string; }) => ({
    accountNumber: account.accountNumber,
    accountType: account.accountType,
    balance: adjustBalance(account.balance, account.debitCredit),
    name: account.name,
    debitCredit: account.debitCredit,
  })); 

  function adjustBalance(balance: number, debitCredit: string): number {
    switch (debitCredit) {
      case 'debit':
        return balance * -1; // credit balance to positive
      case 'credit':
        return balance;
      default:
        return balance; 
    }
  }
}


fetchBudgetFromSPList = async (accountNumbers: number[], month: string, year: string) => {
  if (!this.props.selectedBudgetList) {
    console.error("Selected list ID is not provided");
    return [];
  }

  try {
    const client = await this.props.context.msGraphClientFactory.getClient("3");
    const siteUrl = this.props.context.pageContext.web.serverRelativeUrl;
    const hostname = window.location.hostname;
    const siteId = await client.api(`/sites/${hostname}:${siteUrl}`).get().then(response => response.id);
    const listId = this.props.selectedBudgetList;

    let itemsResponse = await client.api(`/sites/${siteId}/lists/${listId}/items`)
      .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
      .filter(`(fields/M_x00e5_ned eq '${month}' or fields/M_x00e5_ned eq '*') and fields/_x00c5_rstal eq '${year}'`)
      .expand('fields($select=Budget,Kontonr,_x00c5_rstal,M_x00e5_ned)')
      .get();

    const results: { budget: any; budgetYear: any; budgetMonth: any; accountNumber: any; }[] = [];

    itemsResponse.value.forEach((item: { fields: { Budget: any; _x00c5_rstal: any; M_x00e5_ned: any; Kontonr: any; }; }) => {
      const entry = {
        budget: item.fields.Budget,
        budgetYear: item.fields._x00c5_rstal,
        budgetMonth: item.fields.M_x00e5_ned,
        accountNumber: item.fields.Kontonr,
      };
      results.push(entry);
    });

    console.log("Fetched budget items:", results);
    return results;
  } catch (error) {
    console.error("Failed to fetch budget data:", error);
    return [];
  }
}


async fetchAccountDetailsByNumber(accountNumber: number): Promise<AccountDetail | null> {
  const url = `https://restapi.e-conomic.com/accounts?filter=accountNumber$eq:${accountNumber}`;
  
  const response = await fetch(url, {
      method: 'GET',
      headers: {
        'X-AppSecretToken': "1NglExETODyTzOEr5aqNxmhjmE9VjWOli2lhoEcao5g",
        'X-AgreementGrantToken': "CwRfoodzmqD2b1cRsJpF27zwaCILiSNrjlC7JejFtB81",
        'Content-Type': "application/json"
      },
  });

  if (!response.ok) {
      console.error(`Failed to fetch account details for account number: ${accountNumber}`);
      return null;
  }

  const data = await response.json();
  if (data.collection.length === 0) {
      console.error("No account details found for the specified account number.");
      return null;
  }
  const account = data.collection[0];
  return {
      accountNumber: account.accountNumber,
      accountType: account.accountType,
      name: account.name,
      balance: account.balance,
      debitCredit: account.debitCredit,
      budget: 0,
      budgetYear: new Date().getFullYear(),
      budgetMonth: '', 
      regnSum: 'No',
      deviation: 0,
      addSpacer: 'No',
      boldStyle: 'No',
      budgetText: '',
     
  };
}


mathForDeviation = (accounts: AccountDetail[]) => {
  return accounts.map(account => {
    const totalBudget = account.budget + (account.specificMonthBudget || 0);
    if (typeof account.balance === 'number' && typeof totalBudget === 'number') {
      const deviation = account.balance - totalBudget;
      console.log(`Calculating deviation for account ${account.accountNumber}: Balance (${account.balance}) - TotalBudget (${totalBudget}) = Deviation (${deviation})`);
      return {
        ...account,
        deviation: deviation,
      };
    } else {
      return account;
    }
  });
}




calculateBudgetSum = (accounts: AccountDetail[]) => {
  let cumulativeBudgetSum = 0;
  let currentGroup = 0;

  const updatedAccounts = accounts.map((account: AccountDetail, index: number) => {
    console.log(`Processing account ${index}:`, account);
    if (account.regnSum === 'Yes') {
      // End of current iteration, set the budget to the cumulative sum
      account.budget = cumulativeBudgetSum;
      account.group = currentGroup;
      cumulativeBudgetSum = 0; // Reset for the next group
      currentGroup++;
    } else {
      // Add the current account's budget and specificMonthBudget to the cumulative sum
      cumulativeBudgetSum += account.budget + (account.specificMonthBudget || 0);
      account.group = currentGroup;
    }
    console.log(`Processed account ${index}:`, account);
    return account;
  });

  console.log('Updated accounts after budget sum calculation:', updatedAccounts);

  const finalAccounts = updatedAccounts.map((account: AccountDetail, index: number, arr: AccountDetail[]) => {
    console.log(`Calculating final budget for account ${index}:`, account);
    if (account.budget === 0 && index > 0) {
      let sum = 0;
      for (let i = 0; i < index; i++) {
        if (arr[i].group === account.group) {
          sum += arr[i].budget;
        }
      }
      account.budget = sum;
    }
    console.log(`Final budget for account ${index}:`, account);
    return account;
  });

  console.log('Final accounts after processing all budgets:', finalAccounts);
  return finalAccounts;
}


combineAccountInfo = async (monthIndex?: number, currentYear?: string) => {
  try {
    const accountStyles = await this.fetchAccountNumbers();
    const accountNumbers = accountStyles.map(a => a.accountNumber);

    const monthToUse = monthIndex !== undefined ? monthIndex.toString() : (new Date().getMonth() + 1).toString();
    const yearToUse = currentYear !== undefined ? currentYear : new Date().getFullYear().toString();
    const currentMonthName = this.getMonthName(Number(monthToUse));

    const [accountDetails, budgetData] = await Promise.all([
      this.fetchAccountDetails(accountNumbers),
      this.fetchBudgetFromSPList(accountNumbers, monthToUse, yearToUse)
    ]);

    console.log("Account Details:", accountDetails);
    console.log("Budget Data:", budgetData);

    const combinedDetails = accountDetails.map((detail: { accountNumber: number; }) => {
      const styleData = accountStyles.find(style => style.accountNumber === detail.accountNumber);
      const budgetInfo = budgetData.filter(budget => budget.accountNumber === detail.accountNumber);

      let budgetText = '';
      const defaultBudget = budgetInfo.find(budget => budget.budgetMonth === '*');
      const specificMonthBudget = budgetInfo.find(budget => budget.budgetMonth === monthToUse);

      if (defaultBudget) {
        budgetText += `Kr. ${formatNumber(defaultBudget.budget)}`;
      }
      if (specificMonthBudget) {
        if (budgetText) {
          budgetText += '\n';
        }
        budgetText += `Ekstra for ${currentMonthName}: Kr. ${formatNumber(specificMonthBudget.budget)}`;
      }

      return {
        ...detail,
        boldStyle: styleData ? styleData.boldStyle : 'No',
        addSpacer: styleData ? styleData.addSpacer : 'No',
        budget: defaultBudget ? defaultBudget.budget : 0,
        specificMonthBudget: specificMonthBudget ? specificMonthBudget.budget : 0,
        regnSum: styleData ? styleData.regnSum : 'No',
        budgetText: budgetText,
        group: 0 // Initialize group
      };
    });

    console.log("Combined Details before calculating running total and deviation:", combinedDetails);

    const combinedDetailsWithRunningTotal = this.calculateBudgetSum(combinedDetails);
    const accountsWithDeviation = this.mathForDeviation(combinedDetailsWithRunningTotal);

    // Ensure budgetText is set for all accounts, including those with regnSum 'Yes'
    const finalAccounts = accountsWithDeviation.map(account => {
      if (account.regnSum === 'Yes') {
        account.budgetText = `Kr. ${formatNumber(account.budget)}`;
      }
      return account;
    });

    this.setState({ accounts: finalAccounts });
    console.log("Final combined details with running total and deviation:", finalAccounts);
  } catch (error) {
    console.error("Failed to combine account info:", error);
  }
}


private monthNames: string[] = ["Januar", "Februar", "Marts", "April", "Maj", "Juni", "Juli", "August", "September", "Oktober", "November", "December"];
getMonthName(monthNumber: number): string {
  const index = monthNumber - 1;
  return index >= 0 && index < 12 ? this.monthNames[index] : "All Months";
}


setCurrentMonthFromData(month: string) {
  if (month === "*") {
    this.setState({ currentMonth: "All Months" });
  } else {
    const monthNumber = parseInt(month);
    if (!isNaN(monthNumber) && monthNumber >= 1 && monthNumber <= 12) {
      this.setState({ currentMonth: this.getMonthName(monthNumber) });
    } else {
      console.error("Invalid month number:", month);
      this.setState({ currentMonth: "Invalid Month" });
    }
  }
}


nextMonthButton = () => {
  this.setState(prevState => {
    const currentIndex = this.monthNames.indexOf(prevState.currentMonth);
    const nextIndex = (currentIndex + 1) % this.monthNames.length;
    return { currentMonth: this.monthNames[nextIndex] };
  }, () => {
    this.loadDataForMonth(this.state.currentMonth);
  });
}

prevMonthButton = () => {
  this.setState(prevState => {
    const currentIndex = this.monthNames.indexOf(prevState.currentMonth);
    const prevIndex = currentIndex === 0 ? this.monthNames.length - 1 : currentIndex - 1;
    return { currentMonth: this.monthNames[prevIndex] };
  }, () => {
    this.loadDataForMonth(this.state.currentMonth);
  });
}


loadMonthlyData(month: string) {
  const monthNumber = this.monthNames.indexOf(month) + 1; 
  const currentYear = new Date().getFullYear().toString();
  const accountNumbers = this.state.accounts.map(account => account.accountNumber);

  this.fetchBudgetFromSPList(accountNumbers, monthNumber.toString(), currentYear)
    .then(budgetData => {
      if (budgetData.length === 0 || budgetData.every((item: { budget: undefined; }) => item.budget === undefined)) {
        this.fetchBudgetFromSPList(accountNumbers, "*", currentYear)
          .then(fallbackData => this.processBudgetData(fallbackData));
      } else {
        this.processBudgetData(budgetData);
      }
    })
    .catch(error => {
      console.error("Failed to fetch budget data for month:", error);
    });
}

loadDataForMonth = async (month: string) => {
  const currentYear = new Date().getFullYear().toString() as string; 
  const monthIndex = this.monthNames.indexOf(month) + 1;

  const accountNumbers = this.cachedAccountDetails.map(account => account.accountNumber);

  try {
    const budgetData = await this.fetchBudgetFromSPList(accountNumbers, monthIndex.toString(), currentYear);
    this.processBudgetData(budgetData);
    this.combineAccountInfo(monthIndex, currentYear);
  } catch (error) {
    console.error("Failed to fetch budget data for month:", error);
  }
}


updateAccountsWithBudgetData(budgetData: any[]) {
  const updatedAccounts = this.state.accounts.map(account => {
    const budgetInfo = budgetData.find(b => b.accountNumber === account.accountNumber);
    return budgetInfo ? { ...account, ...budgetInfo } : account;
  });
  this.setState({ accounts: updatedAccounts });
}


fetchDataForMonth = async (month: string) => {
  const monthNumber = this.monthNames.indexOf(month) + 1; // Convert month name to its numerical representation
  const currentYear = new Date().getFullYear().toString(); // Dynamically getting the current year as a string
  const accountNumbers = this.state.accounts.map(account => account.accountNumber);

  try {
    const budgetData = await this.fetchBudgetFromSPList(accountNumbers, monthNumber.toString(), currentYear);
    const fullYearData = await this.fetchBudgetFromSPList(accountNumbers, "*", currentYear);

    // Combine data from specific month and full year
    const updatedAccounts = this.state.accounts.map(account => {
      const specificMonthData = budgetData.find((item: { accountNumber: number; }) => item.accountNumber === account.accountNumber);
      const yearlyData = fullYearData.find((item: { accountNumber: number; }) => item.accountNumber === account.accountNumber);

      return {
        ...account,
        budget: specificMonthData ? specificMonthData.budget : (yearlyData ? yearlyData.budget : account.budget),
        budgetMonth: specificMonthData ? specificMonthData.budgetMonth : (yearlyData ? yearlyData.budgetMonth : month),
      };
    });

    this.setState({ accounts: updatedAccounts });
  } catch (error) {
    console.error("Failed to fetch budget data for month:", error);
  }
}


updateDataForMonth = (month: string) => {
  const accountNumbers = this.state.accounts.map(account => account.accountNumber);
  const monthNumber = this.monthNames.indexOf(month) + 1;
  const currentYear = new Date().getFullYear().toString();

  this.fetchBudgetFromSPList(accountNumbers, monthNumber.toString(), currentYear)
    .then(budgetData => {
      const updatedAccounts = this.state.accounts.map(account => {
        const newBudgetData = budgetData.find((item: { accountNumber: number; budgetMonth: string; }) => item.accountNumber === account.accountNumber && item.budgetMonth === month);
        return newBudgetData ? {...account, budget: newBudgetData.budget} : account;
      });

      this.setState({ accounts: updatedAccounts });
    })
    .catch(error => {
      console.error("Failed to fetch budget data for month:", error);
    });
}


handleMonthChange = (selectedMonth: string) => {
  this.setState({ currentMonth: selectedMonth }, () => {
    this.updateDataForMonth(this.state.currentMonth);
  });
}


getStyleForAccount = (account: AccountDetail) => {
  return account.boldStyle === 'Yes' ? { fontWeight: 'bold' } : {};
};


getAddSpacerProp = (account: AccountDetail) => {
  if (account.addSpacer === 'Yes') {
    return (
      <FluentUI.TableRow key={`spacer-${account.accountNumber}`}>
        <FluentUI.TableCell colSpan={4}className={styles.tableHeaderRow}></FluentUI.TableCell>
      </FluentUI.TableRow>
    );
  }
  return null;
};


accountExists(accountNumber: number): boolean {
  const exists = this.cachedAccountDetails.some(account => account.accountNumber === accountNumber);
  console.log(`Check if account ${accountNumber} exists: ${exists}`);
  return exists;
}


createNewKonti = async () => {
  const { formAccountNumber, kontiOptions } = this.state;
  const accountNumber = formAccountNumber ? parseInt(formAccountNumber) : null;

  if (accountNumber === null || isNaN(accountNumber)) {
    alert('Indsæt konti gyldigt nummer');
    return;
  }

  const accountDetails = await this.fetchAccountDetailsByNumber(accountNumber);
  if (!accountDetails) {
    alert("Failed to fetch account details. Please check the account number and try again.");
    return;
  }

  const toggleState = kontiOptions === 'yes';

  try {
    const client = await this.props.context.msGraphClientFactory.getClient("3");
    const siteUrl = this.props.context.pageContext.web.serverRelativeUrl;
    const hostname = window.location.hostname;
    const siteId = await client.api(`/sites/${hostname}:${siteUrl}`).get().then(response => response.id);
    const kontiListId = this.props.selectedKontiList;

    await client.api(`/sites/${siteId}/lists/${kontiListId}/items`).post({
      fields: {
        Kontonavn: accountDetails.name,
        Kontotype: accountDetails.accountType,
        Kontonr: accountNumber,
        BoldStyle: toggleState ? "Yes" : "No",
        AddSpacer: toggleState ? "Yes" : "No",
        RegnSum: toggleState ? "Yes" : "No"
      }
    });

    this.setState(prevState => ({
      isDialogVisibleKonti: false,
      accounts: [...prevState.accounts, accountDetails]
    }));

    alert('Konti blev tilføjet!');
  } catch (error) {
    console.error("Fejlede under tilføjelse af konti:", error);
    alert(`Fejlede under tilføjelse af konti: ${error.message}`);
  }
};


deleteKonti = async () => {
  const { kontiToDelete } = this.state;

  if (kontiToDelete === null) {
    return;
  }

  try {
    const client = await this.props.context.msGraphClientFactory.getClient("3");
    const siteUrl = this.props.context.pageContext.web.serverRelativeUrl;
    const hostname = window.location.hostname;
    const siteId = await client.api(`/sites/${hostname}:${siteUrl}`).get().then(response => response.id);
    const kontiListId = this.props.selectedKontiList;

    const itemsResponse = await client.api(`/sites/${siteId}/lists/${kontiListId}/items`)
      .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
      .filter(`fields/Kontonr eq ${kontiToDelete}`)
      .get();

    if (itemsResponse.value.length === 0) {
      throw new Error('The specified list item was not found');
    }

    const itemId = itemsResponse.value[0].id;

    await client.api(`/sites/${siteId}/lists/${kontiListId}/items/${itemId}`).delete();

    this.setState(prevState => ({
      isDialogVisibleDeleteKonti: false,
      kontiToDelete: null,
      accounts: prevState.accounts.filter(account => account.accountNumber !== kontiToDelete)
    }));

    alert('Konti blev slettet!');
  } catch (error) {
    console.error("Fejlede under sletning af konti:", error);
    alert(`Fejlede under sletning af konti: ${error.message}`);
  }
};


async updateKontiList(listId: string, item: any, siteId: string, client: any) {
  await client.api(`/sites/${siteId}/lists/${listId}/items`)
    .post({
      fields: item
    });
}


handleCreateBudget = async () => { // ekstra metode der blev beholdt for testing. laver bøde budget og konti oprettelse i en.
  const { formAccountNumber, formYear, formMonth, formBudget } = this.state;
  const accountNumber = formAccountNumber ? parseInt(formAccountNumber) : null;
  const year = formYear ? parseInt(formYear) : null;
  const budget = formBudget ? parseFloat(formBudget) : null;

  if (accountNumber === null || isNaN(accountNumber) || 
      year === null || isNaN(year) || 
      !formMonth || 
      budget === null || isNaN(budget)) {
      alert('Please fill in all fields with valid values.');
      return;
  }

  const accountDetails = await this.fetchAccountDetailsByNumber(accountNumber);
  if (!accountDetails) {
      alert("Failed to fetch account details. Please check the account number and try again.");
      return;
  }

  try {
      const client = await this.props.context.msGraphClientFactory.getClient("3");
      const siteUrl = this.props.context.pageContext.web.serverRelativeUrl;
      const hostname = window.location.hostname;
      const siteId = await client.api(`/sites/${hostname}:${siteUrl}`).get().then(response => response.id);
      const budgetListId = this.props.selectedBudgetList;
      const kontiListId = this.props.selectedKontiList;


      await client.api(`/sites/${siteId}/lists/${budgetListId}/items`).post({
          fields: {
              Title: accountDetails.name,
              Kontonr: accountNumber,
              _x00c5_rstal: year,
              M_x00e5_ned: formMonth,
              Budget: budget,
              
          }
      });

      await client.api(`/sites/${siteId}/lists/${kontiListId}/items`).post({
        fields: {
          Kontonavn: accountDetails.name,
          Kontotype: accountDetails.accountType,
          Kontonr: accountNumber,
          BoldStyle: accountDetails.boldStyle || "No",
          AddSpacer: accountDetails.addSpacer || "No",
          RegnSum: "No"
        }
    });

      this.setState(prevState => ({
          isDialogVisibleBudget: false,
          accounts: [...prevState.accounts, accountDetails]
      }));

      alert('Budget created successfully!');
  } catch ( error) {
      console.error("Failed to create budget:", error);
      alert(`Failed to create budget: ${error.message}`);
  }
};


handleEditBudget = async () => {
  const { kontiToEditBudget } = this.state;

  if (kontiToEditBudget === null) {
    return;
  }

  try {
    const client = await this.props.context.msGraphClientFactory.getClient("3");
    const siteUrl = this.props.context.pageContext.web.serverRelativeUrl;
    const hostname = window.location.hostname;
    const siteId = await client.api(`/sites/${hostname}:${siteUrl}`).get().then(response => response.id);
    const budgetListId = this.props.selectedBudgetList;

    const itemsResponse = await client.api(`/sites/${siteId}/lists/${budgetListId}/items`)
      .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
      .filter(`fields/Kontonr eq ${kontiToEditBudget}`)
      .get();

    if (itemsResponse.value.length === 0) {
      throw new Error('The specified list item was not found');
    }


    await client.api(`/sites/${siteId}/lists/${budgetListId}/items`).update({
      fields: {

          
      }
  });

    this.setState(prevState => ({
      isDialogVisibleDeleteKonti: false,
      kontiToDelete: null,
      accounts: prevState.accounts.filter(account => account.accountNumber !== kontiToEditBudget)
    }));

    alert('Budget blev opdateret!');
  } catch (error) {
    console.error("Fejlede under opdatering af budget:", error);
    alert(`Fejlede under opdatering af budget: ${error.message}`);
  }
}

toggleDialogOpretBudget = () => {
  this.setState(prevState => ({ isDialogVisibleBudget: !prevState.isDialogVisibleBudget }));
};

toggleDialogEditBudget = (event?: React.MouseEvent<HTMLButtonElement>) => {
  const accountNumber = this.state.kontiToEditBudget;
  this.setState(prevState => ({
    isDialogVisibleEditBudget: !prevState.isDialogVisibleEditBudget,
    kontiToEditBudget: accountNumber !== undefined ? accountNumber : prevState.kontiToEditBudget
  }));
};

openEditBudgetDialog = (accountNumber: number) => {
  this.setState({
    isDialogVisibleEditBudget: true,
    kontiToEditBudget: accountNumber
  });
};


toggleDialogOpretKonti = () => {
  this.setState(prevState => ({
    isDialogVisibleKonti: !prevState.isDialogVisibleKonti,
    kontiOptions: 'standard'
  }));
};


toggleDialogDeleteKonti = (event?: React.MouseEvent<HTMLButtonElement>) => {
  const accountNumber = this.state.kontiToDelete;
  this.setState(prevState => ({
    isDialogVisibleDeleteKonti: !prevState.isDialogVisibleDeleteKonti,
    kontiToDelete: accountNumber !== undefined ? accountNumber : prevState.kontiToDelete
  }));
};

openDeleteDialog = (accountNumber: number) => {
  this.setState({
    isDialogVisibleDeleteKonti: true,
    kontiToDelete: accountNumber
  });
};


getComboBoxOptions = (): IComboBoxOption[] => {
  const specialOption: IComboBoxOption = { key: '*', text: 'For hele året' };

  const monthOptions: IComboBoxOption[] = this.monthNames.map((name, index): IComboBoxOption => ({
      key: index + 1, 
      text: name
  }));

  return [...monthOptions, specialOption];
}


onRenderOption = (option: IComboBoxOption): JSX.Element => {
  return (
    <span style={{ fontWeight: option.key === '*' ? 'bold' : 'normal' }}>
      {option.text}
    </span>
  );
};

handleMouseEnter = (accountNumber: number) => {
  this.setState({ hoveredRow: accountNumber });
};

handleMouseLeave = (accountNumber: number) => {
  this.setState({ hoveredRow: accountNumber });
};


public render(): React.ReactElement<IBudgetDashProps> {
  const { hasTeamsContext } = this.props;
  const { accounts, currentMonth, isDialogVisibleBudget, isDialogVisibleKonti, isDialogVisibleEditBudget, isDialogVisibleDeleteKonti, loading } = this.state;

  if (loading) {
    return <div>Indlæser...</div>;
  }

  return (
    <section className={`${styles.budgetDash} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.cardReplacement}>
        <Toolbar className={styles.toolbarStyle}>
          <div className={styles.toolbarMiddle}>
            <ToolbarButton onClick={this.prevMonthButton} aria-label="Forrige" icon={<ArrowPreviousRegular />} />
            <div className={styles.displayBudgetPeriod}>
              <Title3 className={styles.titleStyle}>Budget for {currentMonth}</Title3>
            </div>
            <ToolbarButton onClick={this.nextMonthButton} aria-label="Næste" icon={<ArrowNextRegular />} />
            <Toggle
              label="Budget visning for året"
            />
          </div>

          <Dialog
  hidden={!isDialogVisibleKonti}
  onDismiss={this.toggleDialogOpretKonti}
  modalProps={{ isBlocking: false }}
  dialogContentProps={{
    type: DialogType.largeHeader,
    title: 'Tilføj ny konti til budget',
    subText: 'Tilføj valid konti nummer.'
  }}
>
  <TextField
    label="Konti nummer"
    value={this.state.formAccountNumber}
    onChange={(_, newValue) => this.setState({ formAccountNumber: newValue })}
    required
  />
  <Toggle
    label="Sum konti"
    onText="Ja"
    offText="Nej"
    checked={this.state.kontiOptions === 'yes'}
    onChange={(_, checked) => this.setState({ kontiOptions: checked ? 'yes' : 'standard' })}
  />
  <DialogFooter>
    <PrimaryButton onClick={this.createNewKonti} text="Tilføj" />
    <DefaultButton onClick={this.toggleDialogOpretKonti} text="Annullere" />
  </DialogFooter>
</Dialog>


          <Dialog
            hidden={!isDialogVisibleEditBudget}
            onDismiss={this.toggleDialogEditBudget}
            modalProps={{ isBlocking: false }}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: `Redigere budget for konti ${this.state.formAccountNumber}`,
            }}
          >
            <TextField
              label={`Redigere budget for konti${this.state.formAccountNumber}`}
              value={this.state.formBudget}
              onChange={(_, newValue) => this.setState({ formBudget: newValue })}
              required
            />
            <ComboBox
              label="Måned"
              placeholder="Vælg en måned at opdatere budget for"
              options={this.getComboBoxOptions()}
              onChange={(_, option?: IComboBoxOption) => this.setState({ formMonth: option ? option.key.toString() : undefined })}
              required
              onRenderOption={this.onRenderOption}
            />
            <DialogFooter>
              <PrimaryButton onClick={this.handleEditBudget} text="Opdater budget" />
              <DefaultButton onClick={() => this.toggleDialogEditBudget()} text="Annuller" />
            </DialogFooter>
          </Dialog>

          <Dialog
            hidden={!isDialogVisibleBudget}
            onDismiss={this.toggleDialogOpretBudget}
            modalProps={{ isBlocking: false }}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: 'Opret nyt budget',
              subText: 'Udfyld detaljer om nyt budget.'
            }}
          >
            <TextField
              label="Konti nummer"
              value={this.state.formAccountNumber}
              onChange={(_, newValue) => this.setState({ formAccountNumber: newValue })}
              required
            />
            <TextField
              label="Årstal"
              value={this.state.formYear}
              onChange={(_, newValue) => this.setState({ formYear: newValue })}
              required
            />
            <ComboBox
              label="Måned"
              placeholder="Vælg en måned for budget"
              options={this.getComboBoxOptions()}
              onChange={(_, option?: IComboBoxOption) => this.setState({ formMonth: option ? option.key.toString() : undefined })}
              required
              onRenderOption={this.onRenderOption}
            />
            <TextField
              label="Budget"
              value={this.state.formBudget}
              onChange={(_, newValue) => this.setState({ formBudget: newValue })}
              required
            />
            <Toggle
              label="Total sum felt"
              onText="Yes"
              offText="No"
              checked={this.state.regnSum === "Yes"}
              onChange={(_, checked) => this.setState({ regnSum: checked ? "Yes" : "No" })}
            />
            <DialogFooter>
              <PrimaryButton onClick={this.handleCreateBudget} text="Opret" />
              <DefaultButton onClick={this.toggleDialogOpretBudget} text="Annullere" />
            </DialogFooter>
          </Dialog>

          <div className={styles.toolbarRight}>
            <div className={styles.menuContainer}>
              <Menu>
                <MenuTrigger>
                  <ToolbarButton aria-label="Menu" icon={<MoreHorizontal24Filled />} />
                </MenuTrigger>
                <MenuPopover className={styles.menuPopOver}>
                  <MenuList>
                    <MenuItem onClick={this.toggleDialogOpretBudget}>Opret budget</MenuItem>
                  </MenuList>
                </MenuPopover>
              </Menu>
            </div>
          </div>
        </Toolbar>
        <div>
        <FluentUI.Table>
            <FluentUI.TableHeader className={styles.tableHeaderRow}>
              <FluentUI.TableRow>
                <FluentUI.TableCell>Konto {<AddSquare16Regular onClick={this.toggleDialogOpretKonti} />}</FluentUI.TableCell>
                <FluentUI.TableCell>Budget</FluentUI.TableCell>
                <FluentUI.TableCell>Opnået</FluentUI.TableCell>
                <FluentUI.TableCell>Afvigelse</FluentUI.TableCell>
              </FluentUI.TableRow>
            </FluentUI.TableHeader>
            <FluentUI.TableBody>
              {accounts.map((account) => (
                account.accountType === AccountType.Heading ? (
                  <FluentUI.TableRow key={account.accountNumber} style={{ fontWeight: 'bold' }}>
                    <FluentUI.TableCell colSpan={4}>{account.name}</FluentUI.TableCell>
                  </FluentUI.TableRow>
                ) : (
                  <>
                    <FluentUI.TableRow
                      key={account.accountNumber}
                      style={this.getStyleForAccount(account)}
                      onMouseEnter={() => this.handleMouseEnter(account.accountNumber)}
                      onMouseLeave={() => this.handleMouseLeave(account.accountNumber)}
                    >
                      <FluentUI.TableCell>
                        <div className={styles.budgetContainer}>
                          {account.name}
                          {this.state.hoveredRow === account.accountNumber && (
                            <Delete16Regular className={styles.editIcon} onClick={() => this.openDeleteDialog(account.accountNumber)} />
                          )}
                        </div>
                      </FluentUI.TableCell>
                      <FluentUI.TableCell>
                        <div className={styles.budgetContainer}>
                          <FluentUI.Textarea
                            value={account.budgetText}
                            readOnly
                            style={{ width: '100%', height: 'auto', resize: 'none', background: 'transparent', border: 'none' }}
                          />
                          {this.state.hoveredRow === account.accountNumber && (
                            <Edit16Regular className={styles.editIcon} onClick={() => this.openEditBudgetDialog(account.accountNumber)} />
                            )}
                        </div>
                      </FluentUI.TableCell>
                      <FluentUI.TableCell>{`Kr. ${formatNumber(account.balance)}`}</FluentUI.TableCell>
                      <FluentUI.TableCell>{`Kr. ${formatNumber(account.deviation)}`}</FluentUI.TableCell>
                    </FluentUI.TableRow>
                    {this.getAddSpacerProp(account)}
                  </>
                )
              ))}
            </FluentUI.TableBody>
          </FluentUI.Table>
        </div>
      </div>

      <Dialog
        hidden={!isDialogVisibleDeleteKonti}
        onDismiss={() => this.toggleDialogDeleteKonti()}
        modalProps={{ isBlocking: false }}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Bekræft sletning',
          subText: 'Er du sikker på, at du vil slette denne konti?'
        }}
      >
        <DialogFooter>
          <PrimaryButton onClick={this.deleteKonti} text="Slet" />
          <DefaultButton onClick={() => this.toggleDialogDeleteKonti()} text="Annuller" />
        </DialogFooter>
      </Dialog>
    </section>
  );
}




}