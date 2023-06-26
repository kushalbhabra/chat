const inputString = "This is a sample text [doc1] containing [doc4] citations [doc3].";
const citationRegex = /\[(doc([0-9]))\]/g;
const citationIndexes = [];

const replacedString = inputString.replace(citationRegex, function(match, p1, p2) {
  const citationIndex = parseInt(p2);
  citationIndexes.push(citationIndex);
  return `^(${citationIndex})`;
});

console.log(replacedString);
console.log(citationIndexes);
