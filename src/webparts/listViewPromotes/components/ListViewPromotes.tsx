/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-unused-expressions */
/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/no-explicit-any */

import * as React from 'react';
import styles from './ListViewPromotes.module.scss';
import type { IListViewPromotesProps } from './IListViewPromotesProps';
import { IListViewPromotesState } from './IListViewPromotesState';
import SharePointService from '../../../services/SharePoint/spService';
import { escape } from '@microsoft/sp-lodash-subset';
import { css } from "office-ui-fabric-react";
import { IListItem } from '../../../services/SharePoint/IListItem';
import { Icon } from "@fluentui/react";

export default class ListViewPromotes extends React.Component<IListViewPromotesProps, IListViewPromotesState> {
  constructor(props: IListViewPromotesProps) {
    super(props);

    // bind methods
    this.getItems = this.getItems.bind(this);
    this.handleSortChange = this.handleSortChange.bind(this);
    this.handleSortOrderChange = this.handleSortOrderChange.bind(this);
    this.sortItems = this.sortItems.bind(this);
    this.setFilters = this.setFilters.bind(this);
    this._onChange = this._onChange.bind(this);
    this.filterArray = this.filterArray.bind(this);
    this.clearFilters = this.clearFilters.bind(this);
    this.closeFilters = this.closeFilters.bind(this);
    this.openFilters = this.openFilters.bind(this);
    this.handlePageChange = this.handlePageChange.bind(this);

    // set initial state
    this.state = {
      items: [],
      filtereditems: [],
      loading: false,
      error: 'null',
      sortBy: this.props.employeeLastName ? this.props.employeeLastName.split('+')[0] : this.props.employeeName ? this.props.employeeName.split('+')[0] : '',
      sortOrder: 'asc',
      filterBy: [],
      activefilters: [],
      currentPage: 1,
      itemsPerPage: 30
    }
  }

  public handleSortChange(event: React.ChangeEvent<HTMLSelectElement>):void {
    this.setState({
      sortBy: event.target.value
    })

    this.sortItems(event.target.value, this.state.sortOrder)
  }

  public clearFilters():void {
    this.setState({
      activefilters: [],
      filtereditems: this.state.items
    })

    const checkboxes: NodeListOf<HTMLInputElement> = document.querySelectorAll('input[data-dropdown="btn-dropdown-checkbox-element"]');
    const checkboxArray: HTMLInputElement[] = Array.from(checkboxes)
    checkboxArray.map((checkbox: { checked: boolean; }) => { checkbox.checked = false; });
  }

  public closeFilters():void {
    const myElement = document.getElementById('filterMenu');
    myElement?.classList.remove(styles.active);
  }

  public openFilters():void {
    const myElement = document.getElementById('filterMenu');
    myElement?.classList.add(styles.active);
  }

  private _handleClick(field: string, isActive: boolean):void {
    const updatedArray = this.state.filterBy.map(item => 
      item.field === field ? { ...item, isActive: isActive === true ? false : true } : { ...item, isActive: false }
    )

    this.setState({
      filterBy: updatedArray
    })
  }

  public handleSortOrderChange(sortOrder:string):any {
    this.setState({
      sortOrder: this.state.sortOrder === 'asc' ? 'desc' : 'asc'
    })

    this.sortItems(this.state.sortBy, sortOrder)
  }

  public sortItems(sortBy:string, sortOrder:string):void {
    const itemsToSort = this.state.filtereditems;

    itemsToSort.sort((a, b) => {
      const aValue = a[sortBy].toLowerCase().replace(/\W/g, '');
      const bValue = b[sortBy].toLowerCase().replace(/\W/g, '');

      if (aValue < bValue) {
        return sortOrder === 'asc' ? -1 : 1;
      } else if (aValue > bValue) {
        return sortOrder === 'asc' ? 1 : -1;
      } else {
        return 0;
      }
    });

    this.setState({
      filtereditems: itemsToSort,
      currentPage: 1 // Reset to first page after sorting
    })
  }

  private _onChange(field:string, option: string) {  

    // first check if field exists at all
    const array = this.state.activefilters
    const indexoffield = array.findIndex((item) => item.field === field)

    // if indexoffield = -1 this means it doesn't exist so we just add the field object and add the option
    if(indexoffield === -1) {
      const obj = [...this.state.activefilters, {
          field: field,
          options: [option]
        }
      ]

      // Apply the filter function
      const filteredArray = this.filterArray(this.state.items, obj);

      filteredArray.sort((a, b) => {
        const aValue = a[this.state.sortBy];
        const bValue = b[this.state.sortBy];
  
        if (aValue < bValue) {
          return this.state.sortOrder === 'asc' ? -1 : 1;
        } else if (aValue > bValue) {
          return this.state.sortOrder === 'asc' ? 1 : -1;
        } else {
          return 0;
        }
      });

      this.setState({
        activefilters: obj,
        filtereditems: filteredArray,
        currentPage: 1 // Reset to first page after filtering
      })
    }
    // else it exists so now we have to check to see if the option exists
    else {
      const optionarray = array[indexoffield].options
      let result:string[] = []

      // iterate through array and add/subtract option
      if (optionarray.includes(option)) {
        // Item exists, remove it
        result = optionarray.filter(e => e !== option)
      } else {
        // Item doesn't exist, add it
        result = optionarray.concat(option);
      }

      let finalfilter

      // if result is empty remove it from state items.filter(item => item.id !== id)
      if(!result.length) {
        finalfilter = this.state.activefilters.filter(item => item.field !== field)
      }
      // else if array contains option - remove it - else add it
      else{
        array[indexoffield].options = result

        finalfilter = array
      }

      // Apply the filter function
      const filteredArray = this.filterArray(this.state.items, finalfilter);

      filteredArray.sort((a, b) => {
        const aValue = a[this.state.sortBy];
        const bValue = b[this.state.sortBy];
  
        if (aValue < bValue) {
          return this.state.sortOrder === 'asc' ? -1 : 1;
        } else if (aValue > bValue) {
          return this.state.sortOrder === 'asc' ? 1 : -1;
        } else {
          return 0;
        }
      });

      this.setState({
        activefilters: finalfilter,
        filtereditems: filteredArray,
        currentPage: 1 // Reset to first page after filtering
      })
    }
  }

  private filterArray(array1: any[], array2: any[]):IListItem[] {
      // Convert array2 into an object that maps fields to their allowed options
      const filterCriteria = array2.reduce((acc: { [x: string]: any; }, { field, options }: any) => {
          acc[field] = options;
          return acc;
      }, {});
  
      // Filter array1 based on the criteria
      return array1.filter((item: { [x: string]: any; }) => {
          // Check if the item satisfies the "OR" condition for each field, AND across fields
          return Object.keys(filterCriteria).every(field => {
              // Check if the field's value is one of the allowed options
              return filterCriteria[field].includes(item[field]);
          });
      });
  }

  public handlePageChange(event: React.MouseEvent<HTMLButtonElement>, page: number): void {
    this.setState({
      currentPage: page
    });
  }

  public render(): React.ReactElement<IListViewPromotesProps> {
    const {
      webpartTitle,
      description,
      employeeName,
      imageURL,
      newCareerLevel,
      solutionFunction,
      region,
      country,
      office,
      peopleSearchURL,
      showoverview,
      overview,
      sortFields,
      filterFields,
    } = this.props;

    const FilterIcon = () => <Icon iconName="FilterSolid" />;
    const CloseIcon = () => <Icon iconName="ErrorBadge" />
    const cssClasses = css(styles.employeeCard, showoverview && styles.withOverview);

    const {
      currentPage,
      itemsPerPage,
      filtereditems 
    } = this.state;

    const indexOfLastItem = currentPage * itemsPerPage;
    const indexOfFirstItem = indexOfLastItem - itemsPerPage;
    const currentItems = filtereditems.slice(indexOfFirstItem, indexOfLastItem);
    const pageNumbers = [];

    for (let i = 1; i <= Math.ceil(filtereditems.length / itemsPerPage); i++) {
      pageNumbers.push(i);
    }

    const renderPageNumbers = () => {
      const maxPagesToShow = 5;
      const totalPages = pageNumbers.length;
      let startPage = Math.max(1, currentPage - Math.floor(maxPagesToShow / 2));
      const endPage = Math.min(totalPages, startPage + maxPagesToShow - 1);

      if (endPage - startPage < maxPagesToShow - 1) {
        startPage = Math.max(1, endPage - maxPagesToShow + 1);
      }

      const pages = [];
      if (startPage > 1) {
        pages.push(<button key={1} onClick={(e) => this.handlePageChange(e, 1)} className={currentPage === 1 ? styles.active : ''}>1</button>);

        if (startPage > 2) {
          pages.push(<span key="ellipsis-start">...</span>);
        }
      }

      for (let i = startPage; i <= endPage; i++) {
        pages.push(
          <button key={i} onClick={(e) => this.handlePageChange(e, i)} className={currentPage === i ? styles.active : ''} disabled={currentPage === i}>
            {i}
          </button>
        );
      }

      if (endPage < totalPages) {
        if (endPage < totalPages - 1) {
          pages.push(<span key="ellipsis-end">...</span>);
        }

        pages.push(<button key={totalPages} onClick={(e) => this.handlePageChange(e, totalPages)} className={currentPage === totalPages ? styles.active : ''}>{totalPages}</button>);
      }

      return pages;
    };

    return (
      <section className={`${styles.listViewPromotes}`}>
        {webpartTitle && <h2 className={styles.webparttitle}>{webpartTitle}</h2>}
        {description && <p className={styles.description}>{escape(description)}</p>}

        {this.state.items && <><button onClick={() => this.openFilters()} className={styles.openFilters}><FilterIcon /> Open Filters</button><div className={styles.filteringsortingcontainer} id="filterMenu">
          <button onClick={() => this.closeFilters()} className={styles.closeFilters}><CloseIcon /> Close Filters</button>
          <div className={styles.filteringcontainer}>
            {filterFields && <>{this.state.filterBy.map((item, index) => <div key={index} className={styles.dropdownGroup} data-dropdown="btn-dropdown-element">
                  <button onClick={() => this._handleClick(item.field, item.isActive)} className={item.isActive ? styles.active : ''} key={index}>
                    {item.displayname}
                    { this.state.activefilters.some(obj => obj.field === item.field) && <FilterIcon /> }
                  </button>
                  <div className={styles.dropdown}>
                    <div className={styles.dropdownscroll}>
                      {item.options.sort().map((option, index) => {
                        return (
                          <div key={index} className={styles.dropdownOption}>
                            <label>
                              <input type='checkbox' name={item.field} value={option} onChange={() => this._onChange(item.field, option)} data-dropdown="btn-dropdown-checkbox-element" />
                              {option}
                            </label>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </div>
              )} <button onClick={() => this.clearFilters()} className={styles.clearBtn}>Clear Filters</button></>
            }
          </div>
          {sortFields &&
            <div className={styles.sortingcontainer}>
              <div className={styles.selectContainer}>
                <label id="sortinglabel" className={styles.sortinglabel}>Sorting by:</label>
                <select aria-labelledby="sortinglabel" value={this.state.sortBy} onChange={this.handleSortChange} className={styles.sorting}>
                  {sortFields.map((item, index) => <option key={index} value={item.split('+')[0]}>{item.split('+')[1]}</option>
                  )}
                </select>
              </div>

              <label className={styles.switch}>
                <input type="checkbox" onClick={() => this.handleSortOrderChange(this.state.sortOrder === 'asc' ? 'desc' : 'asc')} />
                <span className={styles.slider}>
                  <span className={styles.ascText}><Icon iconName="Up" /></span>
                  <span className={styles.descText}><Icon iconName="Down" /></span>
                </span>
              </label>
            </div>
          }
          </div><div className={styles.employeeList}>
            {currentItems.length > 0 ? currentItems.map((item, index) => <div key={index} className={cssClasses}>
              <div className={styles.inner}>
                {imageURL && <div className={styles.photo}><img src={item[imageURL.split('+')[0]].Url} alt={item[employeeName.split('+')[0]]} /></div>}
                <div className={styles.employeeInfo}>
                  <div className={styles.name}>{item[peopleSearchURL.split('+')[0]] ? <a href={item[peopleSearchURL.split('+')[0]].Url} className={styles.namelink} target="_blank" rel="noreferrer">{item[employeeName.split('+')[0]]}</a> : item[employeeName.split('+')[0]]}</div>
                  {item[newCareerLevel.split('+')[0]] && <div className={styles.level}>{item[newCareerLevel.split('+')[0]]}</div>}
                  {item[solutionFunction.split('+')[0]] && <div>{item[solutionFunction.split('+')[0]]}</div>}
                  {item[region.split('+')[0]] && <div>{item[region.split('+')[0]]}</div>}
                  {item[country.split('+')[0]] && <div>{item[country.split('+')[0]]}</div>}
                  {item[country.split('+')[0]] && <div>{item[office.split('+')[0]]}</div>}
                  {showoverview && item[overview.split('+')[0]] && <div className={styles.overview}>{item[overview.split('+')[0]]}</div>}
                </div>
              </div>
            </div>
            ) : <div className={styles.noResults}>There are no employees that match the filters you have chosen.</div>}
          </div>
          <div className={styles.pagination}>
            <button onClick={(e) => this.handlePageChange(e, 1)} disabled={currentPage === 1}>First</button>
            <button onClick={(e) => this.handlePageChange(e, currentPage - 1)} disabled={currentPage === 1}>Prev</button>
              {renderPageNumbers()}
            <button onClick={(e) => this.handlePageChange(e, currentPage + 1)} disabled={currentPage === pageNumbers.length}>Next</button>
            <button onClick={(e) => this.handlePageChange(e, pageNumbers.length)} disabled={currentPage === pageNumbers.length}>Last</button>
          </div>
          <div className={styles.resultsCount}>
            Showing {indexOfFirstItem + 1} to {Math.min(indexOfLastItem, filtereditems.length)} of {filtereditems.length} results
          </div>
          </>
        }
      </section>
    );
  }

  public componentDidMount(): void {

    const handleClickOutside = (event: { target: any; }) => {
      const dropdownRef = document.querySelectorAll('[data-dropdown="btn-dropdown-element"]');
      const parentContainer = event.target.closest('[data-dropdown="btn-dropdown-element"]');

      if (dropdownRef && parentContainer && !Array.from(dropdownRef).includes(parentContainer)) {
        const updatedArray = this.state.filterBy.map(item => ({
          ...item,
          isActive: false
        }));
    
        this.setState({
          filterBy: updatedArray
        })
      }
    };

    document.addEventListener('mousedown', handleClickOutside);  

    if(this.props.listId){
      this.getItems();
    }
  }

  public componentDidUpdate(prevProps: Readonly<IListViewPromotesProps> & Readonly<{ children?: React.ReactNode; }>): void {
    if (prevProps !== this.props) {
      this.getItems();
    }
  }

  public setFilters(items:IListItem[]):void {
    const {
      filterFields
    } = this.props;

    filterFields && filterFields.map((item, index) =>
      {
        const filteroptions = items.map(a => a[item.split('+')[0]]).filter((value, index, self) => self.indexOf(value) === index)
        const obj = [...this.state.filterBy, {
            field: item.split('+')[0],
            isActive: false,
            displayname: item.split('+')[1],
            options: filteroptions
          }
        ]

        this.setState({
          filterBy: obj
        });
      }
    )
  }

  public getItems(): void {

    this.setState({
      loading: true
    });

    const {
      listId,
      employeeName,
      employeeLastName,
      imageURL,
      newCareerLevel,
      solutionFunction,
      region,
      country,
      office,
      promotionLevel,
      peopleSearchURL,
      overview
    } = this.props;

    const selectedFields = [];

    employeeName && selectedFields.push(employeeName.split('+')[0]);
    employeeLastName && selectedFields.push(employeeLastName.split('+')[0]);
    imageURL && selectedFields.push(imageURL.split('+')[0]);
    newCareerLevel && selectedFields.push(newCareerLevel.split('+')[0]);
    solutionFunction && selectedFields.push(solutionFunction.split('+')[0]);
    region && selectedFields.push(region.split('+')[0]);
    country && selectedFields.push(country.split('+')[0]);
    office && selectedFields.push(office.split('+')[0]);
    promotionLevel && selectedFields.push(promotionLevel.split('+')[0]);
    peopleSearchURL && selectedFields.push(peopleSearchURL.split('+')[0]);
    overview && selectedFields.push(overview.split('+')[0]);
    
    SharePointService.getListItems(listId, selectedFields, employeeLastName ? `&$top=5000&$orderby= ${employeeLastName.split('+')[0]} asc` : employeeName ? `&$top=2000&$orderby= ${employeeName.split('+')[0]} asc` : '&$top=2000').then(items => {

      this.setState({
        filterBy: []
      })
      this.setFilters(items.value);

      this.setState({
        items: items.value,
        filtereditems: items.value,
        loading: false,
        error: 'null'
      });

    }).catch(error => {
        this.setState({
            error: 'Something went wrong!',
            loading: false
        });
    });
  }
}
