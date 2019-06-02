import Mustache from 'mustache';

/**
 * Apply the template to each object in the list
 * @param {string} template A string with mustache templating
 * @param {object[]} jsonData A list of objects containing the data to be used by mustache
 * @returns {object[]}
 */
export const produceDataFromTemplate = (template, jsonData) => {
	return jsonData.map((row) => {
		row.template = Mustache.render(template, row);
		return row;
	});
};
