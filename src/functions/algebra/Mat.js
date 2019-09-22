class Mat {
  /**
   *
   * @param {*[][]} matrix
   * @return {*[]}
   */
  static matrixToVector (matrix) {
    const arr = []
    for (let row of matrix) for (let x of row) arr.push(x)
    return arr
  }

}

export {
  Mat
}