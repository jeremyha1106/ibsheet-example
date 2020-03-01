export const createData = count => {
  var company = ['Google', 'Apple', 'Samsung', 'LG', 'Yahoo', 'Microsoft', 'Metanet', 'SK', 'McDonald', 'Amazon'];
  var country = ['Korea', 'USA', 'Vietnam', 'China', 'France', 'Japan', 'Singapore', 'Thailand', 'Cambodia', 'Taiwan'];

  const data = [];
  for (let i = 0; i < count; i++) {
    data.push({
      "sCompany": company[Math.floor(Math.random() * 10)],
      "sCountry": country[Math.floor(Math.random() * 10)],
      "sSaleQuantity": Math.floor(Math.random() * 100000),
      "sSaleIncrease": Math.floor(Math.random() * 10000),
      "sPrice": Math.floor(Math.random() * 10000000),
      "sSatisfaction": Math.floor(Math.random() * (100 - 50 + 1) + 50),
    });
  }

  return data;
}