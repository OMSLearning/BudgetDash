export const formatNumber = (value: number): string => {
    return new Intl.NumberFormat('da-DK', {
      style: 'decimal',
      maximumFractionDigits: 2, 
    }).format(value);
  };
  