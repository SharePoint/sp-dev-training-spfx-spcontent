# DEMO: Using SPHttpClient to talk to SharePoint

> section 1
> _________
> run generator for solution SPFxContent, react webpart SPFxContent
> create a new list and put 10 items in it
  > Countries: United States, United Kingdom, Russia, China, France, Spain, Germany, Japan, India, Australia, Egypt, Brazil, Chile
  > create interface to list items
    ICountryListItem[]: ID, Title
> implement UX
  > Update the properties on the react component to accept a collection of items
  > create UI for showing items
    > update SASS file with new classes for `.list` & `.item`
    > update UX with <UL> for showing list items
> create method that gets top 10 items upon pushing button & updates it
  > button in react component
    > when clicked, raises event
    > event handled in web part
    > web part does work and re-sets collection of items property on the react component
> _________
> section 2
> _________
> add buttons to create, update & delete
> same as above... when clicked, raises event
> event handled in web part
> web part does work and resets collection
> _________
> section 3
> _________
> update the events that handle getting data from live service... change to use mock service when local
