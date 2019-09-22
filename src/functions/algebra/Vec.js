class Vec {
  /**
   *
   * @param {*[]} arr
   * @return {*[]}
   */
  static distinct (arr) {
    return [...new Set(arr)]
  }
}

export {
  Vec
}