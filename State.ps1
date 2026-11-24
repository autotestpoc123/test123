 initFilterOptions(users?: (UserCod | UserGenpop)[]) {
    if (!users) {
      users = this._userArray;
    }
    if (users) {
      _.each(users, user => {
        if (user.City) {
          this.cityList.add(user.City)
        }
        if (user.Building) {
          this.buildingList.add(user.Building)
        }
      });
      this.cityList = this.sortAlphabetically(this.cityList);
      this.buildingList = this.sortAlphabetically(this.buildingList);
    }
  }

  private sortAlphabetically(values: Set<string>): Set<string> {
    return new Set([...values].sort((first, second) =>
      first.localeCompare(second, 'en', { sensitivity: 'base' })
    ));
  }
