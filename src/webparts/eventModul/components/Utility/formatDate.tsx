export const formatDate = (date: Date | string | number | undefined): string => {
  let newDate: Date = new Date();
  if (date instanceof Date && !isNaN(date.getTime())) {
    newDate = date;
  } else if (date !== null && date !== undefined && typeof date !== 'object') {
    newDate = new Date(date);
  }
  return newDate.getDate() + "." + (newDate.getMonth()+1) + "." + newDate.getFullYear();
};

export const formatDateMonthAndYear = (date: Date | string | number | undefined): string => {
  if (date instanceof Date && !isNaN(date.getTime())) {
    let datestring = "";
    switch (date.getMonth()) {
      case 0:
        datestring = `januar ${date.getFullYear()}`;
        break;
      case 1:
        datestring = `februar ${date.getFullYear()}`;
        break;
      case 2:
        datestring = `marts ${date.getFullYear()}`;
        break;
      case 3:
        datestring = `april ${date.getFullYear()}`;
        break;
      case 4:
        datestring = `maj ${date.getFullYear()}`;
        break;
      case 5:
        datestring = `juni ${date.getFullYear()}`;
        break;
      case 6:
        datestring = `juli ${date.getFullYear()}`;
        break;
      case 7:
        datestring = `august ${date.getFullYear()}`;
        break;
      case 8:
        datestring = `september ${date.getFullYear()}`;
        break;
      case 9:
        datestring = `oktober ${date.getFullYear()}`;
        break;
      case 10:
        datestring = `november ${date.getFullYear()}`;
        break;
      case 11:
        datestring = `december ${date.getFullYear()}`;
        break;
      default:
        datestring = `Ukendt måned ${date.getFullYear()}`;
        break;
    }
    
    return datestring;
  } else {
    return "ukendt dato";
  }
};

export const formatDateMonth = (date: Date | string | number | undefined): string => {
  if (date instanceof Date && !isNaN(date.getTime())) {
    let datestring = "";
    switch (date.getMonth()) {
      case 0:
        datestring = `januar`;
        break;
      case 1:
        datestring = `februar`;
        break;
      case 2:
        datestring = `marts`;
        break;
      case 3:
        datestring = `april`;
        break;
      case 4:
        datestring = `maj`;
        break;
      case 5:
        datestring = `juni`;
        break;
      case 6:
        datestring = `juli`;
        break;
      case 7:
        datestring = `august`;
        break;
      case 8:
        datestring = `september`;
        break;
      case 9:
        datestring = `oktober`;
        break;
      case 10:
        datestring = `november`;
        break;
      case 11:
        datestring = `december`;
        break;
      default:
        datestring = `Ukendt måned`;
        break;
    }
    
    return datestring;
  } else {
    return "ukendt dato";
  }
};

export const renderAmount = (amount: number): string => {
  const formattedAmountDK = new Intl.NumberFormat("da-DK", {
    style: "currency",
    currency: "DKK",
  }).format(amount);

  return formattedAmountDK;
};

export const renderTwoDecimalNumber = (amount: number): string => {
  return amount.toFixed(2);
};


export const renderHours = (decimalHours: number): string => {
  return decimalHours.toLocaleString("da-DK", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

  /** hvis de skal vises som 12:15, så brug nedenstående */
  // const hours = Math.floor(decimalHours);
  // const minutes = Math.round((decimalHours - hours) * 60);
  // const formattedMinutes = minutes < 10 ? '0' + minutes : minutes;
  // return `${hours}:${formattedMinutes}`;
};
