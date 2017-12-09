
# Heroes of Pymoli Data Analysis
- 15-24 years olds make up ~65% of all players
- While 20-24 year olds make up 40% of the players and spent the most money on items in the game, 30-34 year olds spent the most on average per item
- Women not only spent less than any other gender, they also paid less per item, whereas Other / Non-Disclosed spent the most per person.


```python
# Import modules
import pandas as pd
```


```python
# Create a reference the JSON file desired
filename = 'data/purchase_data.json'

# Read the JSON into a Pandas DataFrame
purchaseData = pd.read_json(filename)
purchaseData.head()
```




<div>
<style>
    .dataframe thead tr:only-child th {
        text-align: right;
    }

    .dataframe thead th {
        text-align: left;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Age</th>
      <th>Gender</th>
      <th>Item ID</th>
      <th>Item Name</th>
      <th>Price</th>
      <th>SN</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>38</td>
      <td>Male</td>
      <td>165</td>
      <td>Bone Crushing Silver Skewer</td>
      <td>3.37</td>
      <td>Aelalis34</td>
    </tr>
    <tr>
      <th>1</th>
      <td>21</td>
      <td>Male</td>
      <td>119</td>
      <td>Stormbringer, Dark Blade of Ending Misery</td>
      <td>2.32</td>
      <td>Eolo46</td>
    </tr>
    <tr>
      <th>2</th>
      <td>34</td>
      <td>Male</td>
      <td>174</td>
      <td>Primitive Blade</td>
      <td>2.46</td>
      <td>Assastnya25</td>
    </tr>
    <tr>
      <th>3</th>
      <td>21</td>
      <td>Male</td>
      <td>92</td>
      <td>Final Critic</td>
      <td>1.36</td>
      <td>Pheusrical25</td>
    </tr>
    <tr>
      <th>4</th>
      <td>23</td>
      <td>Male</td>
      <td>63</td>
      <td>Stormfury Mace</td>
      <td>1.27</td>
      <td>Aela59</td>
    </tr>
  </tbody>
</table>
</div>



## Player Count


```python
# Total number of unique players
totalSNs = purchaseData.drop_duplicates(['SN'], keep='first')

# Create DataFrame for Output
playerCount = [{'Total Players':len(totalSNs)}]
playerCount = pd.DataFrame(playerCount)
playerCount
```




<div>
<style>
    .dataframe thead tr:only-child th {
        text-align: right;
    }

    .dataframe thead th {
        text-align: left;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Total Players</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>573</td>
    </tr>
  </tbody>
</table>
</div>



## Purchasing Analysis (Total)


```python
# Number of Unique Items
uniqueItems = len(purchaseData['Item ID'].unique())

# Average Price
avgPrice = purchaseData['Price'].mean()

# Number of purchases
purchases = purchaseData['Item Name'].value_counts()
totalPurchases = purchases.sum()

# Total Revenue
revenue = purchaseData['Price'].sum()

# Create DataFrame for Output
purchasingAnalysis = [{'Number of Unique Items':uniqueItems,
                       'Average Price':avgPrice,
                       'Number of Purchases':totalPurchases,
                       'Total Revenue':revenue}]
purchasingAnalysis = pd.DataFrame(purchasingAnalysis)

# Style output
purchasingAnalysis['Average Price'] = purchasingAnalysis['Average Price'].map('$ {:,.2f}'.format)
purchasingAnalysis['Total Revenue'] = purchasingAnalysis['Total Revenue'].map('$ {:,.2f}'.format)
purchasingAnalysis = purchasingAnalysis[['Number of Unique Items','Average Price','Number of Purchases','Total Revenue']]
# Print results
purchasingAnalysis
```




<div>
<style>
    .dataframe thead tr:only-child th {
        text-align: right;
    }

    .dataframe thead th {
        text-align: left;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Number of Unique Items</th>
      <th>Average Price</th>
      <th>Number of Purchases</th>
      <th>Total Revenue</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>183</td>
      <td>$ 2.93</td>
      <td>780</td>
      <td>$ 2,286.33</td>
    </tr>
  </tbody>
</table>
</div>



## Gender Demographics


```python
# Figure how many player there are per gender
genderDemo = totalSNs['Gender'].value_counts()

# Create a gender Demographics DataFrame
genderDemo = pd.DataFrame(genderDemo)

# Calculate Percentage of players
genderDemo['Percentage of Players'] = genderDemo['Gender'] / len(totalSNs)

# Style Output
genderDemo.index.names = ['Gender']
genderDemo = genderDemo[['Percentage of Players','Gender']]
genderDemo = genderDemo.rename(columns={'Gender':'Total Count'})
genderDemo['Percentage of Players'] = genderDemo['Percentage of Players'].map('{:.2%}'.format)

# Print Results
genderDemo
```




<div>
<style>
    .dataframe thead tr:only-child th {
        text-align: right;
    }

    .dataframe thead th {
        text-align: left;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Percentage of Players</th>
      <th>Total Count</th>
    </tr>
    <tr>
      <th>Gender</th>
      <th></th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>Male</th>
      <td>81.15%</td>
      <td>465</td>
    </tr>
    <tr>
      <th>Female</th>
      <td>17.45%</td>
      <td>100</td>
    </tr>
    <tr>
      <th>Other / Non-Disclosed</th>
      <td>1.40%</td>
      <td>8</td>
    </tr>
  </tbody>
</table>
</div>



## Purchasing Analysis (Gender)


```python
# Groupby gender
genderPurch = purchaseData.groupby(['Gender'])

# Purchase Count
genderPurchTot = genderPurch['Price'].count()

# Average Purchase Price
genderPurchAvg = genderPurch['Price'].mean()

# Total Purchase Price
genderPurchPrice = genderPurch['Price'].sum()

# Normalized Totals
normTotGen = genderPurchPrice/genderDemo['Total Count']

# Create DataFrame for Output
genderPurch = pd.DataFrame({'Purchase Count':genderPurchTot,
                            'Average Purchase Price':genderPurchAvg,
                            'Total Purchase Value':genderPurchPrice,
                            'Normailzed Totals': normTotGen
                           })

# Style output
genderPurch['Average Purchase Price'] = genderPurch['Average Purchase Price'].map('$ {:,.2f}'.format)
genderPurch['Total Purchase Value'] = genderPurch['Total Purchase Value'].map('$ {:,.2f}'.format)
genderPurch['Normailzed Totals'] = genderPurch['Normailzed Totals'].map('$ {:,.2f}'.format)
genderPurch = genderPurch[['Purchase Count','Average Purchase Price','Total Purchase Value','Normailzed Totals']]

# Print Results
genderPurch
```




<div>
<style>
    .dataframe thead tr:only-child th {
        text-align: right;
    }

    .dataframe thead th {
        text-align: left;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Purchase Count</th>
      <th>Average Purchase Price</th>
      <th>Total Purchase Value</th>
      <th>Normailzed Totals</th>
    </tr>
    <tr>
      <th>Gender</th>
      <th></th>
      <th></th>
      <th></th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>Female</th>
      <td>136</td>
      <td>$ 2.82</td>
      <td>$ 382.91</td>
      <td>$ 3.83</td>
    </tr>
    <tr>
      <th>Male</th>
      <td>633</td>
      <td>$ 2.95</td>
      <td>$ 1,867.68</td>
      <td>$ 4.02</td>
    </tr>
    <tr>
      <th>Other / Non-Disclosed</th>
      <td>11</td>
      <td>$ 3.25</td>
      <td>$ 35.74</td>
      <td>$ 4.47</td>
    </tr>
  </tbody>
</table>
</div>



## Purchasing Analysis|


```python
# Create Bins for Age Groups
bins = [0,10,15,20,25,30,35,40,100]
group_names = ['<10','10-14','15-19','20-24','25-29','30-34','35-39','40+']

# Cut data using bins
totalSNs['Age Binned'] = pd.cut(totalSNs.loc[:,'Age'], bins, labels=group_names)

# Group by age bins
ageDemo = totalSNs.groupby(['Age Binned'])

# Create Age Demographics DataFrame
ageDemoNumb = ageDemo['SN'].count()

# Create Age Demo DataFrame
ageDemo = pd.DataFrame({'Player Count':ageDemoNumb})

# Calculate Percentage of players
ageDemo['Percentage of Players'] = ageDemo['Player Count'] / len(totalSNs)

# Style Output
ageDemo.index.names = ['Age']
ageDemo['Percentage of Players'] = ageDemo['Percentage of Players'].map('{:.2%}'.format)

# Print Results
ageDemo
```

    /Users/nupur_mathur/anaconda/lib/python3.6/site-packages/ipykernel_launcher.py:6: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: http://pandas.pydata.org/pandas-docs/stable/indexing.html#indexing-view-versus-copy
      





<div>
<style>
    .dataframe thead tr:only-child th {
        text-align: right;
    }

    .dataframe thead th {
        text-align: left;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Player Count</th>
      <th>Percentage of Players</th>
    </tr>
    <tr>
      <th>Age</th>
      <th></th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>10-14</th>
      <td>54</td>
      <td>9.42%</td>
    </tr>
    <tr>
      <th>15-19</th>
      <td>139</td>
      <td>24.26%</td>
    </tr>
    <tr>
      <th>20-24</th>
      <td>234</td>
      <td>40.84%</td>
    </tr>
    <tr>
      <th>25-29</th>
      <td>52</td>
      <td>9.08%</td>
    </tr>
    <tr>
      <th>30-34</th>
      <td>44</td>
      <td>7.68%</td>
    </tr>
    <tr>
      <th>35-39</th>
      <td>25</td>
      <td>4.36%</td>
    </tr>
    <tr>
      <th>40+</th>
      <td>3</td>
      <td>0.52%</td>
    </tr>
    <tr>
      <th>&lt;10</th>
      <td>22</td>
      <td>3.84%</td>
    </tr>
  </tbody>
</table>
</div>



## Purchasing Analysis (Age)


```python
agePurch = totalSNs.groupby(['Age Binned'])

# Purchase Count
agePurchTot = agePurch['Price'].count()

# Average Purchase Price
agePurchAvg = agePurch['Price'].mean()

# Total Purchase Price
agePurchPrice = agePurch['Price'].sum()

# Normalized Totals
normTotAge = agePurchPrice/ageDemoNumb

# Create DataFrame for Output
agePurch = pd.DataFrame({'Purchase Count':agePurchTot,
                            'Average Purchase Price':agePurchAvg,
                            'Total Purchase Value':agePurchPrice,
                            'Normailzed Totals': normTotAge
                           })

# Style output
agePurch.index.names = ['Age']
agePurch['Average Purchase Price'] = agePurch['Average Purchase Price'].map('$ {:,.2f}'.format)
agePurch['Total Purchase Value'] = agePurch['Total Purchase Value'].map('$ {:,.2f}'.format)
agePurch['Normailzed Totals'] = agePurch['Normailzed Totals'].map('$ {:,.2f}'.format)
agePurch = agePurch[['Purchase Count','Average Purchase Price','Total Purchase Value','Normailzed Totals']]

# Print Results
agePurch
```




<div>
<style>
    .dataframe thead tr:only-child th {
        text-align: right;
    }

    .dataframe thead th {
        text-align: left;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Purchase Count</th>
      <th>Average Purchase Price</th>
      <th>Total Purchase Value</th>
      <th>Normailzed Totals</th>
    </tr>
    <tr>
      <th>Age</th>
      <th></th>
      <th></th>
      <th></th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>10-14</th>
      <td>54</td>
      <td>$ 2.89</td>
      <td>$ 155.94</td>
      <td>$ 2.89</td>
    </tr>
    <tr>
      <th>15-19</th>
      <td>139</td>
      <td>$ 2.90</td>
      <td>$ 403.12</td>
      <td>$ 2.90</td>
    </tr>
    <tr>
      <th>20-24</th>
      <td>234</td>
      <td>$ 2.99</td>
      <td>$ 699.12</td>
      <td>$ 2.99</td>
    </tr>
    <tr>
      <th>25-29</th>
      <td>52</td>
      <td>$ 2.92</td>
      <td>$ 151.65</td>
      <td>$ 2.92</td>
    </tr>
    <tr>
      <th>30-34</th>
      <td>44</td>
      <td>$ 3.33</td>
      <td>$ 146.48</td>
      <td>$ 3.33</td>
    </tr>
    <tr>
      <th>35-39</th>
      <td>25</td>
      <td>$ 2.87</td>
      <td>$ 71.64</td>
      <td>$ 2.87</td>
    </tr>
    <tr>
      <th>40+</th>
      <td>3</td>
      <td>$ 2.88</td>
      <td>$ 8.64</td>
      <td>$ 2.88</td>
    </tr>
    <tr>
      <th>&lt;10</th>
      <td>22</td>
      <td>$ 3.16</td>
      <td>$ 69.50</td>
      <td>$ 3.16</td>
    </tr>
  </tbody>
</table>
</div>



## Top Spenders


```python
# Group by SN
topSpenders = purchaseData.groupby(['SN'])

# Purchase Count
topSpendersCount = topSpenders['Item ID'].count()

# Average Purchase Price
topSpendersAvg = topSpenders['Price'].mean()

# Total Purchase Value
topSpendersTot = topSpenders['Price'].sum()

# Create DataFrame for Output
topSpenders = pd.DataFrame({'Purchase Count':topSpendersCount,
                            'Average Purchase Price':topSpendersAvg,
                            'Total Purchase Value':topSpendersTot
                           })

topSpenders = topSpenders.sort_values(['Total Purchase Value'],ascending=False)

# Style output
topSpenders['Average Purchase Price'] = topSpenders['Average Purchase Price'].map('$ {:,.2f}'.format)
topSpenders['Total Purchase Value'] = topSpenders['Total Purchase Value'].map('$ {:,.2f}'.format)
topSpenders = topSpenders[['Purchase Count','Average Purchase Price','Total Purchase Value']]



# Print Results
topSpenders.head(5)
```




<div>
<style>
    .dataframe thead tr:only-child th {
        text-align: right;
    }

    .dataframe thead th {
        text-align: left;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Purchase Count</th>
      <th>Average Purchase Price</th>
      <th>Total Purchase Value</th>
    </tr>
    <tr>
      <th>SN</th>
      <th></th>
      <th></th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>Undirrala66</th>
      <td>5</td>
      <td>$ 3.41</td>
      <td>$ 17.06</td>
    </tr>
    <tr>
      <th>Saedue76</th>
      <td>4</td>
      <td>$ 3.39</td>
      <td>$ 13.56</td>
    </tr>
    <tr>
      <th>Mindimnya67</th>
      <td>4</td>
      <td>$ 3.18</td>
      <td>$ 12.74</td>
    </tr>
    <tr>
      <th>Haellysu29</th>
      <td>3</td>
      <td>$ 4.24</td>
      <td>$ 12.73</td>
    </tr>
    <tr>
      <th>Eoda93</th>
      <td>3</td>
      <td>$ 3.86</td>
      <td>$ 11.58</td>
    </tr>
  </tbody>
</table>
</div>



## Most Popular Items


```python
# Group by Item ID
popItems = purchaseData.groupby(['Item ID','Item Name'])

# Purchase Count
popItemsCount = popItems['Item ID'].count()

# Item Price
popItemsPrice = popItems['Price'].mean()

# Total Purchase Value
popItemsTot = popItems['Price'].sum()

# Create DataFrame for Output
itemSum = pd.DataFrame({'Purchase Count':popItemsCount,
                         'Item Price':popItemsPrice,
                         'Total Purchase Value':popItemsTot
                       })


popItems = itemSum.sort_values(['Purchase Count'],ascending=False)

# Style output
popItems['Item Price'] = popItems['Item Price'].map('$ {:,.2f}'.format)
popItems['Total Purchase Value'] = popItems['Total Purchase Value'].map('$ {:,.2f}'.format)
popItems = popItems[['Purchase Count','Item Price','Total Purchase Value']]

# Print Results
popItmes5 = popItems.head(5)
popItmes5
```




<div>
<style>
    .dataframe thead tr:only-child th {
        text-align: right;
    }

    .dataframe thead th {
        text-align: left;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th></th>
      <th>Purchase Count</th>
      <th>Item Price</th>
      <th>Total Purchase Value</th>
    </tr>
    <tr>
      <th>Item ID</th>
      <th>Item Name</th>
      <th></th>
      <th></th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>39</th>
      <th>Betrayal, Whisper of Grieving Widows</th>
      <td>11</td>
      <td>$ 2.35</td>
      <td>$ 25.85</td>
    </tr>
    <tr>
      <th>84</th>
      <th>Arcane Gem</th>
      <td>11</td>
      <td>$ 2.23</td>
      <td>$ 24.53</td>
    </tr>
    <tr>
      <th>31</th>
      <th>Trickster</th>
      <td>9</td>
      <td>$ 2.07</td>
      <td>$ 18.63</td>
    </tr>
    <tr>
      <th>175</th>
      <th>Woeful Adamantite Claymore</th>
      <td>9</td>
      <td>$ 1.24</td>
      <td>$ 11.16</td>
    </tr>
    <tr>
      <th>13</th>
      <th>Serenity</th>
      <td>9</td>
      <td>$ 1.49</td>
      <td>$ 13.41</td>
    </tr>
  </tbody>
</table>
</div>



## Most Profitable Items


```python
profItems = itemSum.sort_values(['Total Purchase Value'],ascending=False)

# Style output
profItems['Item Price'] = profItems['Item Price'].map('$ {:,.2f}'.format)
profItems['Total Purchase Value'] = profItems['Total Purchase Value'].map('$ {:,.2f}'.format)
profItems = profItems[['Purchase Count','Item Price','Total Purchase Value']]

# Print Results
profItems5 = profItems.head(5)
profItems5
```




<div>
<style>
    .dataframe thead tr:only-child th {
        text-align: right;
    }

    .dataframe thead th {
        text-align: left;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th></th>
      <th>Purchase Count</th>
      <th>Item Price</th>
      <th>Total Purchase Value</th>
    </tr>
    <tr>
      <th>Item ID</th>
      <th>Item Name</th>
      <th></th>
      <th></th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>34</th>
      <th>Retribution Axe</th>
      <td>9</td>
      <td>$ 4.14</td>
      <td>$ 37.26</td>
    </tr>
    <tr>
      <th>115</th>
      <th>Spectral Diamond Doomblade</th>
      <td>7</td>
      <td>$ 4.25</td>
      <td>$ 29.75</td>
    </tr>
    <tr>
      <th>32</th>
      <th>Orenmir</th>
      <td>6</td>
      <td>$ 4.95</td>
      <td>$ 29.70</td>
    </tr>
    <tr>
      <th>103</th>
      <th>Singed Scalpel</th>
      <td>6</td>
      <td>$ 4.87</td>
      <td>$ 29.22</td>
    </tr>
    <tr>
      <th>107</th>
      <th>Splitter, Foe Of Subtlety</th>
      <td>8</td>
      <td>$ 3.61</td>
      <td>$ 28.88</td>
    </tr>
  </tbody>
</table>
</div>


