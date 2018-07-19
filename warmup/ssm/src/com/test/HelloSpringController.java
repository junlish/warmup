package com.test;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.ModelAndView;

@Controller
public class HelloSpringController {
	String message = "Welcome to Spring MVC!";

	@RequestMapping("/hello")
	public ModelAndView showMessage(
			@RequestParam(value = "name", required = false, defaultValue = "Spring") String name) {

		ModelAndView mv = new ModelAndView("hellospring");
		mv.addObject("message", message);
		mv.addObject("name", name);
		return mv; 
	}
}
