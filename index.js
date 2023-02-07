const runExport = async () => {
  var start = new Date();
  // start a timer - just for performance logging :)

  // reset any
  // add a spinner so the user knows something's going on
  document.getElementById("submit").disabled = true;
  document.getElementById("lds-spinner").style.display = block;
  document.getElementById("errorbox").innerText = "";

  // parse inputs from document
  let pat = document.getElementById("pat").value;
  let portfolio_link = document.getElementById("portfolio").value;

  // validate Personal Access Token
  if ((pat == "") | (portfolio_link == "")) {
    let message =
      "Your Personal Access Token is invalid - please check it and try again";
    postError(message);
    return;
  }

  // Attempt to get portfolio GID from portfolio link
  let portfolio_link_split = portfolio_link.split("/");
  if (portfolio_link_split.length < 3) {
    let message =
      "Your portfolio link is invalid - please check it and try again";
    postError(message);
    return;
  }

  let portfolio_gid = portfolio_link_split[portfolio_link_split.length - 2];

  if (!portfolio_gid | isNaN(portfolio_gid) | (portfolio_gid.length < 3)) {
    let message =
      "Your portfolio link is invalid - please check it and try again";
    postError(message);
    return;
  }

  // test that the personal access token works:

  const httpHeaders = { Authorization: `Bearer ${pat}` };
  let resp = await fetch(`https://app.asana.com/api/1.0/users/me`, {
    headers: httpHeaders,
  });

  if (!resp.ok) {
    let message =
      "Your Personal Access Token is invalid - please check it and try again";
    postError(message);

    return;
  }

  // run the portfolio extract function to get a single array of all projects
  let projects = await extractPortfolio(portfolio_gid, {}, httpHeaders);

  // create a consolidated list of all headers, and dedupe any projects which are in multiple portfolios
  let headerSet = new Set();
  let finalProjects = {};

  // for all projects:
  for (let i = 0; i < projects.length; i++) {
    // check if we've already mapped the project.
    if (projects[i]["Project Id"] in finalProjects) {
      // If so, just add any additional properties:
      finalProjects[projects[i]["Project Id"]] = {
        ...projects[i],
        ...finalProjects[projects[i]["Project Id"]],
      };
    } else {
      // otherwise, create a new project
      finalProjects[projects[i]["Project Id"]] = {
        ...projects[i],
      };
    }

    for (const property in projects[i]) {
      // add all properties (custom fields) to our set of headers
      headerSet.add(property);
    }
  }

  // get the consolidated list of projects
  let finalProjectList = Object.values(finalProjects);

  // create the list of headers
  let csvHeaders = [...headerSet];

  // export the data to CSV and download it:
  exportToCsv(csvHeaders, finalProjectList, "PortfolioExport");

  // log the finish time
  var time = new Date() - start;
  console.log(`finished in ${time} ms`);

  // stop running the spinner, allow for another submit.
  document.getElementById("lds-spinner").style.visibility = "hidden";
  document.getElementById("submit").disabled = false;
  return false;
};

function postError(message) {
  // this function puts a red error text box on the page, allows you to submit the form again

  document.getElementById("errorbox").innerText += message;

  document.getElementById("lds-spinner").style.visibility = "hidden";

  document.getElementById("submit").disabled = false;

  return;
}

async function extractPortfolio(portfolio_gid, presetValues, httpHeaders) {
  // A recursive function that takes in a portfolio GID,
  // any portfolio-level custom fields that should apply to all projects
  // and default headers (including authorization)
  // This function returns a list of all projects under that portfolio,
  // including projects in nested portfolios

  let items = [];
  // get all items from the portfolio
  try {
    items = await getAsanaPortfolioItems(portfolio_gid, {
      headers: httpHeaders,
    });
  } catch (error) {
    console.log(error);
    postError(
      "Something went wrong... inpect the page to view the dev console or wait and try again"
    );
  }

  let projects = [];
  let portfolioPromises = [];

  // for each item:
  for (let i = 0; i < items.length; i++) {
    let item = items[i];
    // get the items custom fields into a flat dictionary:
    let itemFields = flattenCustomFields(item);

    if (item["resource_type"] == "project") {
      // if it's a project, flatten the standard project fields, add its custom fields,
      // and then add it to our list of projects
      let newItem = {
        ...flattenProjectFields(item),
        ...itemFields,
        ...presetValues,
      };
      projects.push(newItem);
    } else if (item["resource_type"] == "portfolio") {
      // if it's a portfolio, run this function recursively.
      portfolioPromises.push(
        extractPortfolio(
          item["gid"],
          { ...itemFields, ...presetValues },
          httpHeaders
        )
      );
    }
  }

  // wait for all the nested portfolio responses to come back, add those projects also to the list
  let portfolioResults = await Promise.all(portfolioPromises);
  projects.push(portfolioResults);

  // the above can create nested arrays, this flattens them out:
  let flatProjects = projects.flat(3);

  return flatProjects;
}

async function getAsanaPortfolioItems(portfolio_gid, headers) {
  // this function uses the asana API to get all the projects and portfolios from a portfolio.

  // max retries for rate limited calls
  let complete = false;
  let retryCounter = 0;
  let maxRetries = 10;

  // while we haven't finished the request, keep trying
  while (retryCounter < maxRetries) {
    // get items from the portfolio with the exact fields we want
    const resp = await fetch(
      `https://app.asana.com/api/1.0/portfolios/${portfolio_gid}/items?opt_fields=name,resource_type,archived,color,created_at,current_status_update.(created_by.name|status_type|created_at|text),notes,modified_at,public,owner.name,start_on,due_on,custom_fields.(name|display_value|type|number_value|datetime_value)`,
      headers
    );
    // if we succeeded, return the results
    if (resp.ok) {
      const results = await resp.json();
      return results["data"];
    }

    // if the error is due to lack of permissions or something thats our fault, just stop
    if (resp.status >= 400 && resp.status != 429 && resp.status != 500) {
      document.getElementById("errorbox").innerHTML +=
        errorCodeMap[resp.status] || "";
      break;
    }

    // back off exponentially in case we're hitting rate limits - wait before retrying
    retryCounter++;
    let wait_time = retryCounter * retryCounter * 120;

    await new Promise((r) => setTimeout(r, wait_time));
  }

  return [];
}

function flattenProjectFields(project) {
  // This function is just for mapping API fields to descriptive reporting headers

  console.log(project);
  newProject = {
    "Project Id": escapeText(project["gid"] || ""),
    "Project Name": escapeText(project["name"] || ""),
    "Project Notes": escapeText(project["notes"] || ""),
    "Project Color": escapeText(project["color"] || ""),
    "Project Created At": project["created_at"] || "",
    "Project Data Archived": project["archived"] || "false",
    "Project Current Status Color":
      project["current_status_update"]?.["status_type"] in statusTextMap
        ? statusTextMap[project["current_status_update"]?.["status_type"]]
        : "",
    "Project Current Status Posted By": escapeText(
      project["current_status_update"]?.["created_by"]?.["name"] || ""
    ),
    "Project Current Status Posted On":
      project["current_status_update"]?.["created_at"] || "",
    "Project Current Status Text": escapeText(
      project["current_status_update"]?.["text"] || ""
    ),
    "Project Modified At": project["modified_at"] || "",
    "Project Name": escapeText(project["name"] || ""),
    "Project Owner Name": escapeText(project["owner"]?.["name"] || ""),
    "Project Public": project["public"] || "",
    "Project Start On": project["start_on"] || "",
    "Project Due On": project["due_on"] || "",
  };
  return newProject;
}

function escapeText(text) {
  // we are wrapping text in doublequote(") text delimiters and escaping any doublequotes

  let newText = text.replace(/"/g, '""');
  return '"' + newText + '"';
}

function flattenCustomFields(object) {
  // this function flattens the API provided custom fields to a simple key-value store

  let flattenedFields = {};
  if ("custom_fields" in object) {
    for (let i = 0; i < object["custom_fields"].length; i++) {
      let field = object["custom_fields"][i];

      if (!!field["display_value"]) {
        if (["multi_enum", "enum", "text", "people"].includes(field["type"])) {
          flattenedFields[field["name"]] = escapeText(
            field["display_value"] || ""
          );
        } else if (field.type == "date") {
          flattenedFields[field["name"]] = field["display_value"];
        } else if (field.type == "number") {
          flattenedFields[field["name"]] = field["number_value"];
        }
      } else {
        flattenedFields[field["name"]] = "";
      }
    }
  }
  return flattenedFields;
}

function exportToCsv(headers, projects, fileName) {
  let csvContent = "";

  // write the header row
  csvContent += headers
    .map((h) => {
      return '"' + h + '"';
    })
    .join(",");
  csvContent += "\n";

  // map each project as a new row
  let projectsCsvData = projects
    .map((project) => {
      let rowString = "";
      rowString += headers
        .map((key) => {
          return project[key] || "";
        })
        .join(",");
      return rowString;
    })
    .join("\n");
  
  // join the content to the headers
  csvContent += projectsCsvData;

  // create the file with the data;
  let blob = new Blob([csvContent], { type: "text/csv" });
  let href = window.URL.createObjectURL(blob);

  // force download:
  let link = document.createElement("a");
  link.setAttribute("href", href);
  link.setAttribute("download", "portfolio_export.csv");
  document.body.appendChild(link);
  link.click();
  window.URL.revokeObjectURL(href);

  return;
}

// mapping api error codes to useful information
const errorCodeMap = {
  400: "Something went wrong with the request - check that your portfolio link looks like https://app.asana.com/0/portfolio/12345/list",
  401: "You're not authorized to get this portfolio - check that you pasted your Personal Access Token correctly and your portfolio link is correct",
  403: "You're not authorized to get this portfolio - check that you pasted your Personal Access Token correctly and your portfolio link is correct",
  404: "We couldn't find that portfolio - check that your portfolio link looks like https://app.asana.com/0/portfolio/12345/list",
};

// mapping project status names to colors
const statusTextMap = {
  on_track: "green",
  at_risk: "yellow",
  off_track: "red",
  on_hold: "blue",
  complete: "complete",
};
