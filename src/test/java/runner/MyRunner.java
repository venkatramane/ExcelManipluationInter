package runner;

import org.junit.runner.RunWith;

import cucumber.api.CucumberOptions;
import cucumber.api.junit.Cucumber;

@RunWith(Cucumber.class)
@CucumberOptions(features="C:\\Users\\VENKATRAMAN\\workspace\\ExcelManipulation\\ExcelManipulation\\src\\main\\java\\feature\\excel_manipulation.feature",
					glue= {"step_definition"},
					monochrome=true,
					format={"pretty","html:test-output","json:json_output/cucumber.json","junit:junit_xml/cucumber.xml"},
					strict=true,
					dryRun=false


		
		)

public class MyRunner {

}
